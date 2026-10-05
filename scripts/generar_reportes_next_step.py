"""
Lee los 8 Excel de metricas Mid/SMB (los mas recientes en Descargas) y genera:

  - Un HTML por vendedor con SUS deals que tienen alerta en la revision de
    Next Step (Where are we? vacio, Next Step vacio/sin fecha/vencido, ECD
    invalido/vencido). Si un vendedor no tiene alertas, NO se genera archivo
    para el.
  - Un HTML general (no por vendedor) con el resumen de Account Book
    Attention del equipo (organizaciones asignadas vs farming activo +
    expansion abierta).

Replica exactamente las reglas de "Calidad de CRM" y "Account book
attention" de Dashboard SMB-MM/Dashboard/Dashboard.html, para que los
correos coincidan con lo que muestra el dashboard.

Ejecutar con: python generar_reportes_next_step.py
"""
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl

from paths import DOWNLOADS, GB_CLAUDE

# El equipo sale de equipo.json, no del codigo. El resto de los scripts
# importa TEAM_OWNERS desde aqui, asi que este sigue siendo el punto unico.
from equipo import (  # noqa: F401
    NOMBRE_PANEL, OWNER_ALIASES, PATRON_REPORTES, TEAM_OWNERS,
)

UMBRAL_SIN_ATENCION_DIAS = 7

CATEGORY_MAP = {
    "orgiq": "orgiq",
    "churnassets": "churn",
    "churn": "churn",
    "farmconver": "farmconver",
    "diq": "diq",
    "deals created": "created",
    "created": "created",
    "oa": "oa",
    "far": "far",
}

ALERT_EXPLANATIONS = {
    "Where are we vacío": "El campo <b>Where are we?</b> está vacío. Agrega un resumen del estatus del deal.",
    "Next Step vacío": "No tiene <b>Next Step</b>. Agrega la próxima acción con fecha (ej: 'Call;08/15/26').",
    "Next Step sin fecha": "El <b>Next Step</b> no tiene una fecha reconocible. Agrégala (ej: 'Call;08/15/26').",
    "Next Step fecha inválida": "La fecha del <b>Next Step</b> no es válida (revisa día/mes/año).",
    "Next Step vencido": "El <b>Next Step</b> está vencido. Actualízalo con la próxima acción y fecha.",
    "ECD formato inválido": "El <b>Expected Close Date</b> tiene un formato inválido. Corrígelo.",
    "ECD vencido": "El <b>Expected Close Date</b> ya pasó. Actualiza la fecha esperada de cierre.",
}


# ---------- utilidades replicadas de Dashboard.html ----------

def clean(v):
    return str(v or "").strip().lower()


def money(v):
    if v is None or v == "":
        return 0.0
    try:
        return float(str(v).replace("$", "").replace(",", ""))
    except ValueError:
        return 0.0


def canonical_owner(v):
    name = str(v or "").strip()
    return OWNER_ALIASES.get(name.lower(), name)


def is_closed_deal(r):
    stage = clean(r.get("Deals Sales Stage"))
    return "closed won" in stage or "closed lost" in stage or stage in ("won", "lost")


def is_open_deal(r):
    return clean(r.get("Deals Sales Stage")) not in ("closed won", "closed lost")


def is_service(r):
    return clean(r.get("Deals New/Renew/ProServ")) == "professionalservices gb"


def is_new_business(r):
    return clean(r.get("Deals New/Renew/ProServ")) == "new"


def is_active_farming(r):
    stage = clean(r.get("Farming Farming Stage")).rstrip(".").strip()
    return stage not in (
        "completed",
        "farming unattended",
        "farming unnatended",
        "completed unattended",
        "completed unnatended",
    )


def find_column(rows, names):
    cols = list(rows[0].keys()) if rows else []
    for name in names:
        for c in cols:
            if clean(c) == clean(name):
                return c
    return ""


def unique_by(rows, field):
    seen = set()
    out = []
    for r in rows:
        key = clean(r.get(field))
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def dedupe_deals_por_id(rows):
    """Colapsa las filas del DIQ que son el MISMO deal.

    vTiger exporta una fila por PRODUCTO del deal, no una por deal: un deal con
    dos productos sale dos veces, identica en todo salvo "Products Product
    Name". Sin esto el reporte listaba el deal repetido -- caso real del
    22-sep-2026: OPT26130 (NEW Add-on Freddy AI FDO - Cognosonline) salio dos
    veces en el reporte de Maria, una por "Freddy AI Copilot" y otra por "New
    Freddy AI Agent". En ese export eran 91 deals con filas de sobra.

    Se conserva la PRIMERA fila de cada deal, pero si esa fila trae el producto
    vacio y otra del mismo deal si lo trae, se copia: asi la alerta "Product
    Name vacio" solo salta cuando el deal de verdad no tiene ningun producto, y
    no segun el orden en que vtiger devolvio las filas.

    Las filas sin Deal ID pasan tal cual: mejor un duplicado que perder un deal
    por no poder identificarlo."""
    id_field = find_column(rows, ["Deals Deal ID", "Deal ID", "Deals Potential No", "Potential No"])
    if not id_field:
        return rows, 0
    product_field = find_column(
        rows, ["Products Product Name", "Product Name", "Products Products Name", "Deals Product Name"]
    )

    por_id = {}
    out = []
    colapsadas = 0
    for r in rows:
        key = clean(r.get(id_field))
        if not key:
            out.append(r)
            continue
        if key not in por_id:
            copia = dict(r)
            por_id[key] = copia
            out.append(copia)
            continue
        colapsadas += 1
        guardada = por_id[key]
        if product_field and not str(guardada.get(product_field, "") or "").strip():
            otro = str(r.get(product_field, "") or "").strip()
            if otro:
                guardada[product_field] = r[product_field]
    return out, colapsadas


def filter_owner_rows(rows, field, owner):
    return [r for r in rows if canonical_owner(r.get(field)) == owner]


def fmt_money(v):
    return "$" + format(round(v), ",")


def fmt_pct(v):
    return "-" if v is None else f"{round(v * 100)}%"


def rate(a, b):
    return (a / b) if b else None


MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
NUMERIC_DATE_RE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4}|\d{2})")
WRITTEN_DATE_RE = re.compile(
    r"\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?"
    r"|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?\s*,?\s*(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*(\d{4}|\d{2})\b",
    re.IGNORECASE,
)


def parse_next_step(value):
    text = str(value or "").strip()
    if not text:
        return {"ok": False, "reason": "Next Step vacío"}
    m = NUMERIC_DATE_RE.search(text)
    w = None if m else WRITTEN_DATE_RE.search(text)
    if not m and not w:
        return {"ok": False, "reason": "Next Step sin fecha"}
    match = m or w
    action = text[: match.start()]
    action = re.sub(r"[;,\s]+$", "", action).strip()
    if not action:
        action = re.split(r"[;,.]", text)[0].strip()
    if m:
        month, day, raw_year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        month, day, raw_year = MONTHS.get(w.group(1)[:3].lower(), 0), int(w.group(2)), int(w.group(3))
    year = raw_year + 2000 if raw_year < 100 else raw_year
    try:
        date = datetime(year, month, day)
    except ValueError:
        return {"ok": False, "reason": "Next Step fecha inválida"}
    return {"ok": True, "action": action, "date": date}


def parse_date_cell(v):
    """Excel ya trae la mayoria de fechas como datetime via openpyxl."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return datetime(v.year, v.month, v.day)
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})", s)
    if m:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


# ---------- carga de archivos ----------

# Columna que solo trae el reporte de farming sin producto. Se detecta por ahi
# y no por el nombre, porque "Metrics Farming sin Producto" contiene "far" y el
# detector por nombre lo confundiria con el reporte FAR.
COL_UNICA_FARMNOPROD = "farming organization"


def leer_encabezado(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        fila = next(ws.iter_rows(values_only=True), ())
        wb.close()
        return [str(c).strip().lower() for c in fila if c is not None]
    except Exception:
        return []


def detect_category(path):
    if COL_UNICA_FARMNOPROD in leer_encabezado(path):
        return "farmnoprod"
    norm = path.name.lower()
    # El reporte "Farming sin Producto" puede exportar vacio (0 orgs, sin
    # encabezado) cuando no hay coincidencias; en ese caso el chequeo de
    # columna de arriba no detecta nada y "farming" cae en el fallback de
    # abajo, donde su substring "far" lo confundiria con la categoria "far".
    if "sin producto" in norm:
        return "farmnoprod"
    m = re.search(r"metricas\s*2025\s*-?\s*([a-z ]+?)(?:_\d|\.xlsx)", norm)
    if m:
        token = m.group(1).strip()
        if token in CATEGORY_MAP:
            return CATEGORY_MAP[token]
    for key, category in CATEGORY_MAP.items():
        if key in norm:
            return category
    return None


def find_latest_files():
    candidates = {}
    for path in DOWNLOADS.glob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        # Con dos paneles en la misma PC los exports de ambos equipos conviven
        # en Descargas, y quedarse con el mas reciente de cada categoria hace
        # que un panel lea los numeros del otro sin que nada falle.
        if PATRON_REPORTES and PATRON_REPORTES.lower() not in path.name.lower():
            continue
        category = detect_category(path)
        if not category:
            continue
        best = candidates.get(category)
        if best is None or path.stat().st_mtime > best.stat().st_mtime:
            candidates[category] = path
    return candidates


def check_stale_files(files, today):
    """Reportes de vTiger en Descargas cuya fecha (de modificacion, o sea
    cuando se descargaron/generaron) es anterior a hoy."""
    stale = []
    for cat, path in files.items():
        file_date = datetime.fromtimestamp(path.stat().st_mtime).date()
        if file_date < today:
            stale.append((cat, path, file_date))
    return stale


def read_xlsx_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows_iter = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(rows_iter, [])]
    rows = []
    for raw in rows_iter:
        if not any(v not in (None, "") for v in raw):
            continue
        rows.append({header[i]: raw[i] for i in range(len(header)) if i < len(raw)})
    wb.close()
    return rows


# ---------- logica de negocio (replica Dashboard.html) ----------

def next_step_quality_rows(diq_rows, validation_date):
    next_step_field = find_column(
        diq_rows,
        ["Deals Next Step", "Deals Next step", "Deals NextStep", "Deals Next Steps", "Next Step", "Next step"],
    )
    where_field = find_column(diq_rows, ["Deals Where are we?", "Where are we?", "Deals Where are we"])
    ecd_field = find_column(
        diq_rows, ["Deals Expected Close Date", "Expected Close Date", "Expected Close", "ECD", "Ecd"]
    )
    product_field = find_column(
        diq_rows, ["Products Product Name", "Product Name", "Products Products Name", "Deals Product Name"]
    )
    service_type_field = find_column(
        diq_rows, ["Deals Professional Service Type", "Professional Service Type", "Service Type"]
    )

    checks = []
    for r in diq_rows:
        owner = canonical_owner(r.get("Deals Assigned To"))
        if owner not in TEAM_OWNERS:
            continue
        if is_closed_deal(r):
            continue

        raw_next = str(r.get(next_step_field, "") or "").strip() if next_step_field else ""
        parsed = parse_next_step(raw_next)
        where = str(r.get(where_field, "") or "").strip() if where_field else ""
        raw_ecd = r.get(ecd_field) if ecd_field else None
        ecd_date = parse_date_cell(raw_ecd) if raw_ecd not in (None, "") else None
        product = str(r.get(product_field, "") or "").strip() if product_field else ""
        service_type = str(r.get(service_type_field, "") or "").strip() if service_type_field else ""
        amount = money(r.get("Deals Amount"))

        alerts = []
        if not where:
            alerts.append("Where are we vacío")
        if not next_step_field:
            alerts.append("Falta columna Next Step")
        if not raw_next:
            alerts.append("Next Step vacío")
        if raw_next and not parsed["ok"]:
            alerts.append(parsed["reason"])
        if parsed["ok"] and parsed["date"] < validation_date:
            alerts.append("Next Step vencido")
        if raw_ecd not in (None, "") and not ecd_date:
            alerts.append("ECD formato inválido")
        if ecd_date and ecd_date < validation_date:
            alerts.append("ECD vencido")
        if not product:
            alerts.append("Product Name vacío")
        if is_service(r) and not service_type:
            alerts.append("Service Type vacío")
        if amount == 0:
            alerts.append("Monto en cero")

        checks.append(
            {
                "owner": owner,
                "deal": r.get("Deals Deal Name", "") or "",
                "organization": (r.get("Deals Organization Name") or r.get("Deals Organization Name ") or ""),
                "stage": r.get("Deals Sales Stage", "") or "",
                "amount": amount,
                "where": where,
                "next_step": raw_next,
                "action": parsed.get("action", ""),
                "next_date": parsed.get("date"),
                "ecd": str(raw_ecd) if raw_ecd not in (None, "") else "",
                "ecd_date": ecd_date,
                "product": product,
                "service_type": service_type,
                "is_service_deal": is_service(r),
                "alerts": alerts,
            }
        )
    checks.sort(key=lambda r: (-len(r["alerts"]), r["owner"], -r["amount"]))
    return checks


def farming_without_product_rows(farmnoprod_rows):
    """Farming activo del equipo sin producto asociado. El reporte ya viene
    filtrado desde vTiger; aun asi se revisa el producto vacio y el dueno, igual
    que farmingProductRows del dashboard."""
    product_field = find_column(
        farmnoprod_rows, ["Products Product Name", "Farming Product", "Product Name"]
    )
    org_field = find_column(
        farmnoprod_rows, ["Farming Organization", "Organizations Organization Name"]
    )
    rows = []
    for r in farmnoprod_rows:
        owner = canonical_owner(r.get("Farming Assigned To"))
        if owner not in TEAM_OWNERS:
            continue
        product = str(r.get(product_field, "") or "").strip() if product_field else ""
        if product:
            continue
        rows.append(
            {
                "owner": owner,
                "name": r.get("Farming Farming Name", "") or "",
                "organization": (r.get(org_field, "") or "") if org_field else "",
                "stage": r.get("Farming Farming Stage", "") or "",
            }
        )
    rows.sort(key=lambda r: (r["owner"], r["organization"]))
    return rows


def farming_next_step_rows(active_farming, validation_date):
    """Alertas de Next Step sobre farming activo, con el mismo parseo de fecha
    que los deals. Farming no tiene 'Where are we?' ni ECD, asi que esas dos
    alertas no aplican."""
    next_step_field = find_column(active_farming, ["Farming Next Step", "Farming Next step"])
    checks = []
    for r in active_farming:
        owner = canonical_owner(r.get("Farming Assigned To"))
        if owner not in TEAM_OWNERS:
            continue
        raw_next = str(r.get(next_step_field, "") or "").strip() if next_step_field else ""
        parsed = parse_next_step(raw_next)

        alerts = []
        if not next_step_field:
            alerts.append("Falta columna Next Step")
        if not raw_next:
            alerts.append("Next Step vacío")
        if raw_next and not parsed["ok"]:
            alerts.append(parsed["reason"])
        if parsed["ok"] and parsed["date"] < validation_date:
            alerts.append("Next Step vencido")

        checks.append(
            {
                "owner": owner,
                "name": r.get("Farming Farming Name", "") or "",
                "organization": r.get("Organizations Organization Name", "") or "",
                "stage": r.get("Farming Farming Stage", "") or "",
                "next_step": raw_next,
                "action": parsed.get("action", ""),
                "next_date": parsed.get("date"),
                "alerts": alerts,
            }
        )
    checks.sort(key=lambda r: (-len(r["alerts"]), r["owner"]))
    return checks


def deals_sin_atencion_rows(diq_rows, hoy, umbral=UMBRAL_SIN_ATENCION_DIAS):
    """Deals abiertos del equipo sin ninguna actividad (Modified Time) hace
    mas del umbral. None si el reporte DIQ no trae esa columna."""
    mod_field = find_column(diq_rows, ["Deals Modified Time", "Modified Time"])
    if not mod_field:
        return None
    rows = []
    for r in diq_rows:
        owner = canonical_owner(r.get("Deals Assigned To"))
        if owner not in TEAM_OWNERS or is_closed_deal(r):
            continue
        raw = r.get(mod_field)
        fecha = parse_date_cell(raw) if raw not in (None, "") else None
        if not fecha:
            continue
        dias = (hoy - fecha).days
        if dias <= umbral:
            continue
        rows.append(
            {
                "owner": owner,
                "deal": r.get("Deals Deal Name", "") or "",
                "organization": r.get("Deals Organization Name") or r.get("Deals Organization Name ") or "",
                "amount": money(r.get("Deals Amount")),
                "modificado": fecha,
                "dias": dias,
                "alerts": [f"Sin atención hace {dias} días"],
            }
        )
    rows.sort(key=lambda r: (-r["dias"], r["owner"]))
    return rows


def farming_sin_atencion_rows(active_farming, hoy, umbral=UMBRAL_SIN_ATENCION_DIAS):
    """Farming activo del equipo sin ninguna actividad (Modified Time) hace
    mas del umbral. None si el reporte FAR no trae esa columna."""
    mod_field = find_column(active_farming, ["Farming Modified Time", "Modified Time"])
    if not mod_field:
        return None
    rows = []
    for r in active_farming:
        owner = canonical_owner(r.get("Farming Assigned To"))
        if owner not in TEAM_OWNERS:
            continue
        raw = r.get(mod_field)
        fecha = parse_date_cell(raw) if raw not in (None, "") else None
        if not fecha:
            continue
        dias = (hoy - fecha).days
        if dias <= umbral:
            continue
        rows.append(
            {
                "owner": owner,
                "name": r.get("Farming Farming Name", "") or "",
                "organization": r.get("Organizations Organization Name", "") or "",
                "stage": r.get("Farming Farming Stage", "") or "",
                "modificado": fecha,
                "dias": dias,
                "alerts": [f"Sin atención hace {dias} días"],
            }
        )
    rows.sort(key=lambda r: (-r["dias"], r["owner"]))
    return rows


# Las alertas de producto se listan aparte de las de Next Step.
PRODUCT_ALERT_KEYS = ("Product Name vacío", "Service Type vacío")


def with_only_alerts(rows, keys):
    salida = []
    for r in rows:
        filtradas = [a for a in r["alerts"] if a in keys]
        if filtradas:
            salida.append({**r, "alerts": filtradas})
    return salida


def without_alerts(rows, keys):
    salida = []
    for r in rows:
        filtradas = [a for a in r["alerts"] if a not in keys]
        if filtradas:
            salida.append({**r, "alerts": filtradas})
    return salida


def account_book_attention_rows(org_rows, active_farming, open_opps):
    rows = []
    for owner in TEAM_OWNERS:
        orgs = unique_by(filter_owner_rows(org_rows, "Organizations Assigned To", owner), "Organizations Organization Name")
        farming = [r for r in active_farming if canonical_owner(r.get("Farming Assigned To")) == owner]
        opps = [r for r in open_opps if canonical_owner(r.get("Deals Assigned To")) == owner]
        attention = len(farming) + len(opps)
        if orgs or attention:
            rows.append(
                {
                    "owner": owner,
                    "orgs": len(orgs),
                    "active_farming": len(farming),
                    "open_opps": len(opps),
                    "attention": attention,
                }
            )
    return rows


# ---------- generacion de HTML (GB Advisors design system) ----------

LOGO_SVG = """<svg width="140" height="30" viewBox="0 0 408 86" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M408 54.1533C408 60.9129 402.302 66 393.027 66C383.155 66 376.53 60.216 376 52.6899H386.667C386.932 55.4077 389.781 57.2195 392.894 57.2195C395.81 57.2195 397.029 55.2977 397.029 53.5556C397.029 47.2838 376 52.2857 376 38C376 31.3798 382.559 26 392.232 26C401.772 26 407.271 31.6678 408 39.3333H397.333C397.002 36.6852 395.081 34.8502 391.901 34.8502C389.251 34.8502 386.971 36.5629 386.971 38.4444C386.971 44.6465 407.801 39.6585 408 54.1533Z" fill="#25235A"/>
<path fill-rule="evenodd" clip-rule="evenodd" d="M352 66V46V45.9717V26H362V28.6802C364.941 26.9755 368.357 26 372 26V36C366.473 36 362 40.4665 362 46V66H352Z" fill="#25235A"/>
<path fill-rule="evenodd" clip-rule="evenodd" d="M348 46C348 57.0457 339.046 66 328 66C316.954 66 308 57.0457 308 46C308 34.9543 316.954 26 328 26C339.046 26 348 34.9543 348 46ZM328 56C333.523 56 338 51.5228 338 46C338 40.4772 333.523 36 328 36C322.477 36 318 40.4772 318 46C318 51.5228 322.477 56 328 56Z" fill="#25235A"/>
<path d="M304 54.1533C304 60.9129 298.302 66 289.027 66C279.155 66 272.53 60.216 272 52.6899H282.667C282.932 55.4077 285.781 57.2195 288.894 57.2195C291.81 57.2195 293.029 55.2977 293.029 53.5556C293.029 47.2838 272 52.2857 272 38C272 31.3798 278.559 26 288.232 26C297.772 26 303.271 31.6678 304 39.3333H293.333C293.002 36.6852 291.081 34.8502 287.901 34.8502C285.251 34.8502 282.971 36.5629 282.971 38.4444C282.971 44.6465 303.801 39.6585 304 54.1533Z" fill="#25235A"/>
<circle cx="261" cy="12" r="6" fill="#25235A"/>
<path d="M266 26H256V66H266V26Z" fill="#25235A"/>
<path d="M237.778 66L252 26H241.361L232 52.6841L222.645 26H212L226.222 66H237.778Z" fill="#25235A"/>
<path fill-rule="evenodd" clip-rule="evenodd" d="M208 6H198V28.6791C194.913 26.7289 191.922 26 188 26C176.954 26 168 34.9543 168 46C168 57.0457 176.954 66 188 66C191.643 66 195.058 65.0261 198 63.3243V66H208V46V6ZM198 46C198 40.4772 193.523 36 188 36C182.477 36 178 40.4772 178 46C178 51.5228 182.477 56 188 56C193.523 56 198 51.5228 198 46Z" fill="#25235A"/>
<path fill-rule="evenodd" clip-rule="evenodd" d="M154 28.5V26H164V66H154V63.5C151.055 65.2041 147.624 66 143.981 66C132.945 66 124 57.0464 124 46C124 34.9537 132.945 26 143.981 26C147.624 26 151.055 26.7959 154 28.5ZM144 56C149.523 56 154 51.5228 154 46C154 40.4772 149.523 36 144 36C138.477 36 134 40.4772 134 46C134 51.5228 138.477 56 144 56Z" fill="#25235A"/>
<path fill-rule="evenodd" clip-rule="evenodd" d="M80 6H70V46V66H80V63.3243C82.9417 65.0261 86.3571 66 90 66C101.046 66 110 57.0457 110 46C110 34.9543 101.046 26 90 26C86.3571 26 82.9417 26.9739 80 28.6757V6ZM80 46C80 51.5228 84.4772 56 90 56C95.5228 56 100 51.5228 100 46C100 40.4772 95.5228 36 90 36C84.4772 36 80 40.4772 80 46Z" fill="#25235A"/>
<path fill-rule="evenodd" clip-rule="evenodd" d="M65.9949 46.1337H65.9975V65.8967H65.9924C65.9939 65.9282 65.9954 65.9589 65.9964 65.9897C65.9971 66.0106 65.9975 66.0316 65.9975 66.053C65.9975 77.0686 57.0451 86 45.9987 86C34.9549 86 26 77.0706 26 66.053C26.0045 66.3535 26.0068 66.5068 26.007 66.5068C26.0071 66.5068 26.0048 66.3403 26 66H36C36 71.5092 40.4769 76 46 76C51.5232 76 55.9981 71.5622 55.9981 66.053C55.9981 66.0001 55.993 65.9471 55.993 65.8941H55.9981V63.3333C53.0569 65.0334 49.6442 66 46 66C34.9562 66 26 56.9648 26 45.9471C26 34.9293 34.9549 26 46.0013 26C49.6435 26 53.0585 26.9712 56 28.6682V26H66V45.9471V46C65.9996 46.0135 65.998 46.0269 65.9975 46.0403C65.9962 46.0712 65.9949 46.1022 65.9949 46.1337ZM35.9994 45.9471C35.9994 51.4559 40.4769 56 46 56C51.5232 56 56.0006 51.4559 55.9981 45.9471C55.9981 40.4381 51.4282 36 46 36C40.5718 36 35.9994 40.4381 35.9994 45.9471Z" fill="#25235A"/>
<path d="M19.4999 0H20.5C22.4885 9.79565 30.2045 17.5115 40 19.4999V20.5C30.2045 22.4885 22.4885 30.2045 20.5 40H19.4999C17.5115 30.2045 9.79565 22.4885 0 20.5V19.4999C9.79565 17.5115 17.5115 9.79565 19.4999 0Z" fill="#EA018B"/>
</svg>"""

STYLE = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  :root {
    --gba-magenta: #EA018B; --gba-magenta-100: #FFEBF7; --gba-magenta-800: #920958;
    --gba-purple-deep: #25235A;
    --gba-ink: #12112C; --gba-50: #F8F9FA; --gba-100: #F1F3F5; --gba-200: #E9ECEF;
    --gba-300: #DEE2E6; --gba-500: #ADB5BD; --gba-700: #495057;
    --radius-lg: 24px; --radius-pill: 999px;
  }
  * { box-sizing: border-box; }
  body { font-family: "Inter", ui-sans-serif, system-ui, Arial, sans-serif; color: var(--gba-ink);
         background: var(--gba-50); margin: 0; padding: 0; -webkit-font-smoothing: antialiased; }
  .wrap { max-width: 860px; margin: 0 auto; padding: 40px 24px; }
  .brand-header { margin-bottom: 32px; }
  .brand-header .logo { margin-bottom: 20px; }
  .brand-header h1 { font-size: 32px; line-height: 1.15; letter-spacing: -0.005em; font-weight: 500; margin: 0 0 12px; }
  .rule { display: block; width: 64px; height: 10px; background: var(--gba-magenta); border-radius: var(--radius-pill); margin-bottom: 12px; }
  .meta { font-size: 12px; line-height: 1.4; letter-spacing: 0.02em; text-transform: uppercase; color: var(--gba-700); }
  .greeting { font-size: 16px; line-height: 1.5; margin: 0 0 20px; color: var(--gba-ink); }
  .card { background: #FFFFFF; border-radius: var(--radius-lg); padding: 8px 24px; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { text-align: left; font-size: 11px; letter-spacing: 0.03em; text-transform: uppercase; color: var(--gba-700);
       background: var(--gba-100); padding: 12px 14px; }
  th:first-child { border-radius: 12px 0 0 12px; }
  th:last-child { border-radius: 0 12px 12px 0; }
  td { padding: 14px; border-bottom: 1px solid var(--gba-200); vertical-align: top; }
  tr:last-child td { border-bottom: none; }
  .deal { font-weight: 500; color: var(--gba-ink); }
  .sub { display: block; font-size: 11.5px; color: var(--gba-700); margin-top: 2px; }
  .chips { display: flex; flex-wrap: wrap; gap: 6px; }
  .chip { display: inline-flex; align-items: center; padding: 4px 10px; border-radius: var(--radius-pill);
          font-size: 11px; font-weight: 500; background: var(--gba-magenta-100); color: var(--gba-magenta-800); white-space: nowrap; }
  .count-chip { background: var(--gba-ink); color: #fff; }
  .footer { margin-top: 24px; font-size: 12px; color: var(--gba-700); }
  h2.bloque { font-size: 22px; font-weight: 500; margin: 32px 0 4px; letter-spacing: -0.005em; }
  h2.bloque::after { content: ""; display: block; width: 48px; height: 8px; background: var(--gba-magenta); border-radius: var(--radius-pill); margin-top: 10px; }
  h2.bloque-sub { font-size: 15px; font-weight: 500; color: var(--gba-700); margin: 20px 0 8px; }
</style>
"""


def esc(v):
    return (
        str(v)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


ALERT_SHORT = {
    "Where are we vacío": "Where are we? vacío",
    "Next Step vacío": "Next Step vacío",
    "Next Step sin fecha": "Next Step sin fecha",
    "Next Step fecha inválida": "Fecha Next Step inválida",
    "Next Step vencido": "Next Step vencido",
    "ECD formato inválido": "ECD formato inválido",
    "ECD vencido": "ECD vencido",
    "Product Name vacío": "Product Name vacío",
    "Service Type vacío": "Service Type vacío",
    "Monto en cero": "Monto en cero",
}


def _tabla(titulo, encabezados, filas_html):
    if not filas_html:
        return ""
    ths = "".join(f"<th>{esc(h)}</th>" for h in encabezados)
    return f"""
    <h2 class="bloque-sub">{esc(titulo)}</h2>
    <div class="card">
      <table>
        <tr>{ths}</tr>
        {''.join(filas_html)}
      </table>
    </div>"""


def _chips(alerts):
    return "".join(f'<span class="chip">{esc(ALERT_SHORT.get(a, a))}</span>' for a in alerts)


def build_vendor_html(owner, secciones, generated_on):
    """secciones: dict con deals_next_step, deals_producto, deals_atencion,
    farming_next_step, farming_producto y farming_atencion. Cada alerta sale
    en su sub-lista; un deal/farming con alertas de varias categorias
    aparece una vez en cada una, no todo junto."""
    first_name = owner.split()[0]
    date_str = generated_on.strftime("%d/%m/%Y")

    # --- Deals: Next Step / Where are we? / ECD ---
    filas = []
    for d in secciones["deals_next_step"]:
        next_date_str = d["next_date"].strftime("%d/%m") if d["next_date"] else "-"
        ecd_str = d["ecd_date"].strftime("%d/%m") if d["ecd_date"] else "-"
        filas.append(
            f"""<tr>
          <td><span class="deal">{esc(d['deal'])}</span><span class="sub">{esc(d['organization'])}</span></td>
          <td>{fmt_money(d['amount'])}</td>
          <td>{esc(d['next_step']) or '(vacío)'}<span class="sub">{next_date_str}</span></td>
          <td>{ecd_str}</td>
          <td><div class="chips">{_chips(d['alerts'])}</div></td>
        </tr>"""
        )
    tabla_deals_ns = _tabla("Alertas de Next Step",
                            ["Deal / Organización", "MRR", "Next Step", "ECD", "Corregir"], filas)

    # --- Deals: producto y tipo de servicio ---
    filas = []
    for d in secciones["deals_producto"]:
        # En deals que no son de servicios el Service Type no aplica: mostrarlo
        # como "(vacío)" hacia pensar que era un error del reporte.
        service = (esc(d["service_type"]) or "(vacío)") if d.get("is_service_deal") else "— (no aplica)"
        filas.append(
            f"""<tr>
          <td><span class="deal">{esc(d['deal'])}</span><span class="sub">{esc(d['organization'])}</span></td>
          <td>{fmt_money(d['amount'])}</td>
          <td>{esc(d['product']) or '(vacío)'}</td>
          <td>{service}</td>
          <td><div class="chips">{_chips(d['alerts'])}</div></td>
        </tr>"""
        )
    tabla_deals_prod = _tabla("Alertas de producto",
                              ["Deal / Organización", "MRR", "Producto", "Service Type", "Corregir"], filas)

    # --- Deals: sin atencion ---
    filas = []
    for d in secciones["deals_atencion"]:
        filas.append(
            f"""<tr>
          <td><span class="deal">{esc(d['deal'])}</span><span class="sub">{esc(d['organization'])}</span></td>
          <td>{fmt_money(d['amount'])}</td>
          <td>{d['modificado']:%d/%m/%Y}</td>
          <td><div class="chips">{_chips(d['alerts'])}</div></td>
        </tr>"""
        )
    tabla_deals_atencion = _tabla("Alertas de atención",
                                  ["Deal / Organización", "MRR", "Última actividad", "Corregir"], filas)

    # --- Farmings: Next Step ---
    filas = []
    for f in secciones["farming_next_step"]:
        next_date_str = f["next_date"].strftime("%d/%m") if f["next_date"] else "-"
        filas.append(
            f"""<tr>
          <td><span class="deal">{esc(f['name'])}</span><span class="sub">{esc(f['organization'])}</span></td>
          <td>{esc(f['stage'])}</td>
          <td>{esc(f['next_step']) or '(vacío)'}<span class="sub">{next_date_str}</span></td>
          <td><div class="chips">{_chips(f['alerts'])}</div></td>
        </tr>"""
        )
    tabla_farm_ns = _tabla("Alertas de Next Step",
                           ["Farming / Organización", "Etapa", "Next Step", "Corregir"], filas)

    # --- Farmings: sin producto ---
    filas = []
    for f in secciones["farming_producto"]:
        filas.append(
            f"""<tr>
          <td><span class="deal">{esc(f['name'])}</span><span class="sub">{esc(f['organization'])}</span></td>
          <td>{esc(f['stage'])}</td>
          <td><div class="chips">{_chips(['Product Name vacío'])}</div></td>
        </tr>"""
        )
    tabla_farm_prod = _tabla("Alertas de producto",
                             ["Farming / Organización", "Etapa", "Corregir"], filas)

    # --- Farmings: sin atencion ---
    filas = []
    for f in secciones["farming_atencion"]:
        filas.append(
            f"""<tr>
          <td><span class="deal">{esc(f['name'])}</span><span class="sub">{esc(f['organization'])}</span></td>
          <td>{esc(f['stage'])}</td>
          <td>{f['modificado']:%d/%m/%Y}</td>
          <td><div class="chips">{_chips(f['alerts'])}</div></td>
        </tr>"""
        )
    tabla_farm_atencion = _tabla("Alertas de atención",
                                 ["Farming / Organización", "Etapa", "Última actividad", "Corregir"], filas)

    bloque_deals = ""
    if tabla_deals_ns or tabla_deals_prod or tabla_deals_atencion:
        bloque_deals = f"""
    <h2 class="bloque">Deals</h2>
    {tabla_deals_ns}
    {tabla_deals_prod}
    {tabla_deals_atencion}"""

    bloque_farming = ""
    if tabla_farm_ns or tabla_farm_prod or tabla_farm_atencion:
        bloque_farming = f"""
    <h2 class="bloque">Farmings</h2>
    {tabla_farm_ns}
    {tabla_farm_prod}
    {tabla_farm_atencion}"""

    total = sum(len(secciones[k]) for k in secciones)
    return f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>Calidad de CRM - {esc(owner)}</title>{STYLE}</head>
<body>
  <div class="wrap">
    <div class="brand-header">
      <div class="logo">{LOGO_SVG}</div>
      <h1>Calidad de CRM</h1>
      <span class="rule"></span>
      <div class="meta">Reporte generado el {date_str} &middot; {total} registro(s) con alerta</div>
    </div>
    <p class="greeting">Hola {esc(first_name)}, esto es lo que hay que corregir en vTiger:</p>
    {bloque_deals}
    {bloque_farming}
    <div class="footer">Generado automáticamente a partir de DIQ, FAR y Farming sin Producto ({date_str}). Cualquier duda, contacta a el responsable comercial.</div>
  </div>
</body></html>"""


def build_general_html(attention_rows, generated_on):
    date_str = generated_on.strftime("%d/%m/%Y")
    total_orgs = sum(r["orgs"] for r in attention_rows)
    total_attention = sum(r["attention"] for r in attention_rows)
    rows_html = "".join(
        f"""<tr>
          <td><span class="deal">{esc(r['owner'])}</span></td>
          <td>{r['orgs']}</td>
          <td>{r['active_farming']}</td>
          <td>{r['open_opps']}</td>
          <td><span class="chip count-chip">{fmt_pct(rate(r['attention'], r['orgs']))}</span></td>
        </tr>"""
        for r in attention_rows
    )
    return f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><title>Account Book Attention - General</title>{STYLE}</head>
<body>
  <div class="wrap">
    <div class="brand-header">
      <div class="logo">{LOGO_SVG}</div>
      <h1>Account book attention</h1>
      <span class="rule"></span>
      <div class="meta">Resumen del equipo Mid/SMB &middot; {date_str}</div>
    </div>
    <p class="greeting">Book attention general del equipo: <strong>{fmt_pct(rate(total_attention, total_orgs))}</strong>
      ({total_attention} de {total_orgs} organizaciones con farming activo o expansión abierta).</p>
    <div class="card">
      <table>
        <tr>
          <th>Vendedor</th><th>Organizaciones</th><th>Farming activos</th>
          <th>Expansión abiertas</th><th>% Book attention</th>
        </tr>
        {rows_html}
      </table>
    </div>
    <div class="footer">Generado automáticamente a partir de ORGIQ, FAR y DIQ ({date_str}).</div>
  </div>
</body></html>"""


def main():
    now = datetime.now()
    files = find_latest_files()

    stale = check_stale_files(files, now.date())
    if stale:
        print("=" * 60)
        print("ALERTA: los reportes de vTiger en Descargas tienen fecha anterior a hoy:")
        for cat, path, file_date in stale:
            print(f"  - [{cat}] {path.name} (fecha: {file_date.strftime('%d/%m/%Y')})")
        print(f"Hoy es {now.strftime('%d/%m/%Y')}. Verifica que hayas descargado los reportes")
        print("actualizados de vTiger antes de generar/enviar los reportes de Next Step.")
        print("=" * 60)

    required = ["orgiq", "oa", "diq", "far", "created", "farmconver", "churn", "farmnoprod"]
    missing = [c for c in required if c not in files]
    if missing:
        print("Faltan archivos para: " + ", ".join(missing))
        print("Encontrados:")
        for cat, path in files.items():
            print(f"  {cat}: {path.name}")
        return

    print("Archivos usados:")
    for cat in required:
        print(f"  {cat}: {files[cat].name}")

    data = {cat: read_xlsx_rows(path) for cat, path in files.items()}

    # El DIQ trae una fila por producto del deal; el reporte razona por deal.
    data["diq"], filas_colapsadas = dedupe_deals_por_id(data["diq"])
    if filas_colapsadas:
        print(f"  (DIQ: {filas_colapsadas} fila(s) de mas por deals con varios "
              f"productos, colapsadas a un deal cada uno)")

    validation_date = datetime(now.year, now.month, now.day)
    active_farming = [r for r in data["far"] if is_active_farming(r)]

    # Categorias de alerta, separadas igual que en el dashboard.
    checks = next_step_quality_rows(data["diq"], validation_date)
    con_alerta = [c for c in checks if c["alerts"]]
    deals_next_step = without_alerts(con_alerta, PRODUCT_ALERT_KEYS)
    deals_producto = with_only_alerts(checks, PRODUCT_ALERT_KEYS)
    farming_checks = farming_next_step_rows(active_farming, validation_date)
    farming_next_step = [f for f in farming_checks if f["alerts"]]
    farming_producto = farming_without_product_rows(data["farmnoprod"])

    deals_atencion = deals_sin_atencion_rows(data["diq"], now)
    if deals_atencion is None:
        print('AVISO: al reporte DIQ le falta la columna "Modified Time"; '
              'no se calculan las alertas de atención de deals.')
        deals_atencion = []
    farming_atencion = farming_sin_atencion_rows(active_farming, now)
    if farming_atencion is None:
        print('AVISO: al reporte FAR le falta la columna "Modified Time"; '
              'no se calculan las alertas de atención de farmings.')
        farming_atencion = []

    CATEGORIAS = ("deals_next_step", "deals_producto", "deals_atencion",
                  "farming_next_step", "farming_producto", "farming_atencion")
    por_vendedor = {o: {k: [] for k in CATEGORIAS} for o in TEAM_OWNERS}
    for clave, filas in (("deals_next_step", deals_next_step),
                         ("deals_producto", deals_producto),
                         ("deals_atencion", deals_atencion),
                         ("farming_next_step", farming_next_step),
                         ("farming_producto", farming_producto),
                         ("farming_atencion", farming_atencion)):
        for fila in filas:
            if fila["owner"] in por_vendedor:
                por_vendedor[fila["owner"]][clave].append(fila)
    open_opps = [
        r for r in data["diq"]
        if is_open_deal(r) and not is_service(r) and not is_new_business(r)
        and canonical_owner(r.get("Deals Assigned To")) in TEAM_OWNERS
    ]
    org_rows_team = [r for r in data["orgiq"] if canonical_owner(r.get("Organizations Assigned To")) in TEAM_OWNERS]
    attention_rows = account_book_attention_rows(org_rows_team, active_farming, open_opps)

    out_dir = GB_CLAUDE / "Emails Vendedores"
    out_dir.mkdir(parents=True, exist_ok=True)
    date_suffix = now.strftime("%d-%m-%Y")

    written = []
    for owner in TEAM_OWNERS:
        secciones = por_vendedor[owner]
        if not any(secciones[k] for k in CATEGORIAS):
            continue
        html = build_vendor_html(owner, secciones, now)
        path = out_dir / f"Next Step - {owner}_{date_suffix}.html"
        path.write_text(html, encoding="utf-8")
        written.append(path.name)

    general_html = build_general_html(attention_rows, now)
    general_path = out_dir / f"Account Book Attention - General_{date_suffix}.html"
    general_path.write_text(general_html, encoding="utf-8")
    written.append(general_path.name)

    # Manifiesto de la corrida: deja constancia exacta de QUE archivos produjo
    # esta generacion y de que tan frescos eran los Excel de origen. Sin esto no
    # se puede distinguir un reporte recien hecho de uno que quedo de una
    # corrida anterior del mismo dia (un vendedor sin alertas no genera archivo,
    # y el viejo se queda en la carpeta).
    fuentes = {cat: {"archivo": p.name,
                     "modificado": datetime.fromtimestamp(p.stat().st_mtime).isoformat()}
               for cat, p in files.items()}
    manifiesto = {
        "generado": now.isoformat(),
        "archivos": written,
        "sin_alertas": [o for o in TEAM_OWNERS
                        if not any(por_vendedor[o][k] for k in CATEGORIAS)],
        "fuentes": fuentes,
    }
    manifest_path = Path(__file__).resolve().parent / "ultima_generacion_next_step.json"
    manifest_path.write_text(json.dumps(manifiesto, indent=2, ensure_ascii=False),
                             encoding="utf-8")

    con_email = sum(1 for o in TEAM_OWNERS if any(por_vendedor[o][k] for k in CATEGORIAS))
    print("")
    print(f"Deals abiertos revisados: {len(checks)}")
    print(f"  con alerta de Next Step: {len(deals_next_step)}")
    print(f"  con alerta de producto:  {len(deals_producto)}")
    print(f"  sin atención:            {len(deals_atencion)}")
    print(f"Farmings activos revisados: {len(farming_checks)}")
    print(f"  con alerta de Next Step: {len(farming_next_step)}")
    print(f"  sin producto:            {len(farming_producto)}")
    print(f"  sin atención:            {len(farming_atencion)}")
    print(f"Vendedores con email generado: {con_email} de {len(TEAM_OWNERS)}")
    sin_alertas = manifiesto["sin_alertas"]
    if sin_alertas:
        print(f"Sin alertas (no se genero archivo): {', '.join(sin_alertas)}")
    print(f"\nCarpeta de salida: {out_dir}")
    for name in written:
        print(f"  - {name}")


if __name__ == "__main__":
    main()
