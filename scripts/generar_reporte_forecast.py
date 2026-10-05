"""
Forecast report for el responsable comercial's team: Upside and Commit for the current or next
month. In English -- this report goes out to Freshworks.

  - Upside: deals in "Negotiation" with an Expected Close Date in the chosen
    month.
  - Commit: deals in "PO Invoiced" or "Accepted Proposal" with an Expected
    Close Date in the chosen month.

Category (CX/EX/Device 42/Freddy-AI/Payment Frequency/Other) and segment
(New Business/Expansion) use the exact same rules as the FW Cadence deck
(FW-Cadence-Generator/build_deck.py): same regexes over Deal Name + Deals
Product, joined by Deal ID against the "FW Cadence IQ" report in Downloads
(which also has Deals Product/Primary Email/Contact Name that DIQ doesn't
have). A deal that isn't in that report (not Freshworks) has no product
category -- it can only be "PF" if it's Payment Frequency, or "Other", and
its contact email/name are left blank.

Reuses the helpers in generar_reportes_next_step.py (DOWNLOADS, TEAM_OWNERS,
canonical_owner, clean, money, fmt_money, parse_date_cell, is_service,
read_xlsx_rows) instead of rewriting them.

Usage:
  python generar_reporte_forecast.py                                  Current month, all categories/segments, GB services included
  python generar_reporte_forecast.py --meses proximo                  Next month
  python generar_reporte_forecast.py --meses 2026-10,2026-12         Two specific months
  python generar_reporte_forecast.py --desde actual --hasta 2026-12  Everything from this month to December
  python generar_reporte_forecast.py --excluir-servicios              Drop ProfessionalServices GB deals
  python generar_reporte_forecast.py --categorias CX,EX                Only CX and EX deals
  python generar_reporte_forecast.py --segmentos "New Business"        Only New Business deals
"""
import argparse
import importlib.util
import json
import re
from datetime import datetime
from pathlib import Path

from paths import FORECAST_DIR as OUT_DIR

SCRIPT_DIR = Path(__file__).resolve().parent
MANIFEST_FILE = SCRIPT_DIR / "ultima_generacion_forecast.json"

TEAM_LABEL = "el responsable comercial's Team"

MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July",
            "August", "September", "October", "November", "December"]


def _load_gen_ns():
    spec = importlib.util.spec_from_file_location(
        "gen_next_step", SCRIPT_DIR / "generar_reportes_next_step.py"
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# ---------- CX/EX category (same as FW-Cadence-Generator/build_deck.py) ----------

D42_RE = re.compile(r"device\s?42|\bd42\b|assets?\s*pack|\bassets?\b", re.I)
AI_RE = re.compile(r"\bbots?\b|bot session|freddy|copilot", re.I)
CX_RE = re.compile(
    r"\bFSAS\b|freshsales\s*suite|\bFDO\b|fd[\s\-/]*omni|freshdesk\s*omni|\bomni\b|\bFSA\b"
    r"|freshsales|\bFCH\b|fchat|freshchat|\bFCA\b|fcaller|freshcaller|\bFM\b|freshmarketer"
    r"|\bFD\b|freshdesk", re.I,
)
EX_RE = re.compile(r"\bFS\b|freshservice", re.I)

CATEGORY_OPTIONS = ["CX", "EX", "D42", "AI", "PF", "Other"]
CAT_LABELS = {"CX": "CX", "EX": "EX", "D42": "Device 42", "AI": "Freddy/AI",
             "PF": "Payment Frequency", "Other": "Other"}
CAT_CLASS = {"CX": "ct-cx", "EX": "ct-ex", "D42": "ct-d42", "AI": "ct-ai",
            "PF": "ct-pf", "Other": "ct-ot"}


def clasificar_categoria(deal_name, producto, tipo):
    """Same logic as cats() in build_deck.py: Deal Name + Deals Product (if
    the deal isn't Freshworks, producto comes in empty and it can only end
    up as PF or Other)."""
    texto = f"{deal_name} {producto}"
    c = set()
    if D42_RE.search(texto):
        c.add("D42")
    if AI_RE.search(texto):
        c.add("AI")
    if tipo == "Payment Frequency":
        c.add("PF")
        return sorted(c)
    if CX_RE.search(texto) and not EX_RE.search(deal_name):
        c.add("CX")
    if EX_RE.search(texto):
        c.add("EX")
    if "CX" in c and "EX" in c:
        if CX_RE.search(deal_name) and not EX_RE.search(deal_name):
            c.discard("EX")
        elif EX_RE.search(deal_name) and not CX_RE.search(deal_name):
            c.discard("CX")
    if not (c & {"CX", "EX", "D42", "AI"}):
        c.add("Other")
    return sorted(c)


def find_fw_cadence_deals(gen_ns):
    """Latest 'FW Cadence IQ_*.xlsx' in the vTiger reports folder (NOT
    'FW Cadence IQ 2', which is the installed-base/assets report and isn't
    needed here). Reuses gen_ns.DOWNLOADS instead of hardcoding the folder
    a second time, so both always point to the same place."""
    candidatos = [
        p for p in gen_ns.DOWNLOADS.glob("FW Cadence IQ_*.xlsx")
        if not p.name.startswith("~$")
    ]
    if not candidatos:
        return None
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def fw_info_por_deal_id(gen_ns, fw_path):
    """Deal ID -> {producto, email, contacto, pais} from the FW Cadence IQ
    report (the only one with Deals Product / Primary Email / Contact Name /
    Billing Country -- DIQ doesn't have a country field)."""
    filas = gen_ns.read_xlsx_rows(fw_path)
    mapa = {}
    for r in filas:
        did = str(r.get("Deals Deal ID") or "").strip()
        if did and did not in mapa:
            mapa[did] = {
                "producto": r.get("Deals Product") or "",
                "email": r.get("Deals Primary Email") or "",
                "contacto": r.get("Deals Contact Name") or "",
                "pais": r.get("Organizations Billing Country") or "",
            }
    return mapa


# ---------- target month(s) ----------

MONTH_OPTIONS = ["actual", "proximo"]
MONTH_LABELS = {"actual": "Current month", "proximo": "Next month"}


def mes_a_anio_mes(clave, hoy):
    """Acepta las dos claves relativas de siempre ("actual", "proximo") y
    tambien un mes absoluto "YYYY-MM".

    Lo absoluto se agrego el 16-sep-2026 a pedido de el responsable comercial: el forecast solo
    dejaba elegir este mes y el proximo, y necesitaba ventanas mas largas
    (un trimestre, o meses sueltos mas adelante). Las claves relativas se
    conservan porque las usan las llamadas existentes y la ayuda del CLI."""
    clave = str(clave).strip()
    if clave == "proximo":
        return (hoy.year + 1, 1) if hoy.month == 12 else (hoy.year, hoy.month + 1)
    if clave == "actual":
        return (hoy.year, hoy.month)
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", clave)
    if not m:
        raise ValueError(
            f"Month '{clave}' not understood. Use 'actual', 'proximo' or 'YYYY-MM'.")
    anio, mes = int(m.group(1)), int(m.group(2))
    if not 1 <= mes <= 12:
        raise ValueError(f"Month '{clave}' is out of range (1-12).")
    return (anio, mes)


def rango_de_meses(desde, hasta, hoy):
    """Todos los meses entre desde y hasta, ambos incluidos. Cada extremo
    puede ser relativo o absoluto. Si vienen al reves, se ordenan solos en
    vez de devolver vacio, que desde el panel seria un error silencioso."""
    a = mes_a_anio_mes(desde, hoy)
    b = mes_a_anio_mes(hasta, hoy)
    if a > b:
        a, b = b, a
    meses = set()
    anio, mes = a
    while (anio, mes) <= b:
        meses.add((anio, mes))
        anio, mes = (anio + 1, 1) if mes == 12 else (anio, mes + 1)
    return meses


def meses_objetivo(meses_arg, hoy):
    """Un deal entra si su ECD cae en CUALQUIERA de los meses elegidos --
    se pueden pedir varios a la vez, consecutivos o sueltos."""
    return {mes_a_anio_mes(clave, hoy) for clave in meses_arg}


# ---------- Upside / Commit calculation ----------

STAGE_UPSIDE = {"negotiation"}
STAGE_COMMIT = {"po invoiced", "accepted proposal"}

# Same criterion as "seg" in build_deck.py: Add-On/Cross-Sell/Payment
# Frequency are Expansion, everything else is New Business. This is what
# lets Payment Frequency be isolated as its own filter (it falls under
# Expansion).
TIPOS_EXPANSION = {"new-add-on", "new-cross-sell", "payment frequency"}
SEGMENT_OPTIONS = ["New Business", "Expansion"]


def clasificar_segmento(gen_ns, tipo):
    return "Expansion" if gen_ns.clean(tipo) in TIPOS_EXPANSION else "New Business"


def calcular_forecast(gen_ns, diq_rows, fw_map, meses_destino, incluir_servicios,
                      categorias_permitidas, segmentos_permitidos, reps_permitidos,
                      paises_permitidos):
    upside, commit = [], []
    vistos = set()
    for r in diq_rows:
        owner = gen_ns.canonical_owner(r.get("Deals Assigned To"))
        if owner not in gen_ns.TEAM_OWNERS:
            continue
        if owner not in reps_permitidos:
            continue
        stage_raw = r.get("Deals Sales Stage") or ""
        stage = gen_ns.clean(stage_raw)
        if stage not in STAGE_UPSIDE and stage not in STAGE_COMMIT:
            continue
        ecd = gen_ns.parse_date_cell(r.get("Deals Expected Close Date"))
        if not (ecd and (ecd.year, ecd.month) in meses_destino):
            continue
        if not incluir_servicios and gen_ns.is_service(r):
            continue

        deal_id = str(r.get("Deals Deal ID") or "").strip()
        # DIQ has "Products Product Name": a deal with several products gets
        # one row per product, with the same amount repeated on each. Without
        # this check the deal shows up duplicated in Upside/Commit.
        if deal_id:
            if deal_id in vistos:
                continue
            vistos.add(deal_id)

        deal_name = r.get("Deals Deal Name") or ""
        tipo = r.get("Deals New/Renew/ProServ") or ""
        fw_info = fw_map.get(deal_id, {})
        # "Deals Product" (FW Cadence IQ) viene vacio en la mayoria de los
        # deals -- DIQ trae el mismo dato en "Products Product Name" para
        # esos mismos casos, asi que se combinan las dos fuentes en vez de
        # depender de una sola (sin esto, deals como Teledolar con Freshdesk
        # Omnichannel Pro caian en "Other" en vez de CX).
        producto = f"{fw_info.get('producto', '')} {r.get('Products Product Name', '') or ''}".strip()
        categorias = clasificar_categoria(deal_name, producto, tipo)
        segmento = clasificar_segmento(gen_ns, tipo)
        pais = fw_info.get("pais") or "Unknown"

        if not any(c in categorias_permitidas for c in categorias):
            continue
        if segmento not in segmentos_permitidos:
            continue
        if paises_permitidos is not None and pais not in paises_permitidos:
            continue

        fila = {
            "owner": owner,
            "deal": deal_name,
            "organization": r.get("Deals Organization Name") or r.get("Deals Organization Name ") or "",
            "stage": stage_raw,
            "ecd": ecd,
            "amount": gen_ns.money(r.get("Deals Amount")),
            "is_service": gen_ns.is_service(r),
            "categorias": categorias,
            "segmento": segmento,
            "pais": pais,
            "email": fw_info.get("email", ""),
            "contacto": fw_info.get("contacto", ""),
            "next_step": r.get("Deals Next Step") or "",
            "where": r.get("Deals Where are we?") or "",
        }
        if stage in STAGE_UPSIDE:
            upside.append(fila)
        else:
            commit.append(fila)

    upside.sort(key=lambda f: (-f["amount"], f["owner"]))
    commit.sort(key=lambda f: (-f["amount"], f["owner"]))
    return upside, commit


# ---------- HTML (same visual language as FW-Cadence-Generator/template.html:
# dark gradient cover, KPI cards, elevated tables, category/stage badges with
# the exact same colors as the deck) ----------

def esc(v):
    return (str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# White-on-dark GB Advisors logo (same as LOGO_W in build_deck.py) -- the ink
# logo used elsewhere in this app is near-invisible on the dark cover.
LOGO_SVG_WHITE = (
    '<svg width="140" height="30" viewBox="0 0 408 86" fill="none" xmlns="http://www.w3.org/2000/svg">'
    '<path d="M408 54.1533C408 60.9129 402.302 66 393.027 66C383.155 66 376.53 60.216 376 52.6899H386.667C386.932 55.4077 389.781 57.2195 392.894 57.2195C395.81 57.2195 397.029 55.2977 397.029 53.5556C397.029 47.2838 376 52.2857 376 38C376 31.3798 382.559 26 392.232 26C401.772 26 407.271 31.6678 408 39.3333H397.333C397.002 36.6852 395.081 34.8502 391.901 34.8502C389.251 34.8502 386.971 36.5629 386.971 38.4444C386.971 44.6465 407.801 39.6585 408 54.1533Z" fill="#FFFFFF"/>'
    '<path fill-rule="evenodd" clip-rule="evenodd" d="M352 66V46V45.9717V26H362V28.6802C364.941 26.9755 368.357 26 372 26V36C366.473 36 362 40.4665 362 46V66H352Z" fill="#FFFFFF"/>'
    '<path fill-rule="evenodd" clip-rule="evenodd" d="M348 46C348 57.0457 339.046 66 328 66C316.954 66 308 57.0457 308 46C308 34.9543 316.954 26 328 26C339.046 26 348 34.9543 348 46ZM328 56C333.523 56 338 51.5228 338 46C338 40.4772 333.523 36 328 36C322.477 36 318 40.4772 318 46C318 51.5228 322.477 56 328 56Z" fill="#FFFFFF"/>'
    '<path d="M304 54.1533C304 60.9129 298.302 66 289.027 66C279.155 66 272.53 60.216 272 52.6899H282.667C282.932 55.4077 285.781 57.2195 288.894 57.2195C291.81 57.2195 293.029 55.2977 293.029 53.5556C293.029 47.2838 272 52.2857 272 38C272 31.3798 278.559 26 288.232 26C297.772 26 303.271 31.6678 304 39.3333H293.333C293.002 36.6852 291.081 34.8502 287.901 34.8502C285.251 34.8502 282.971 36.5629 282.971 38.4444C282.971 44.6465 303.801 39.6585 304 54.1533Z" fill="#FFFFFF"/>'
    '<circle cx="261" cy="12" r="6" fill="#FFFFFF"/>'
    '<path d="M266 26H256V66H266V26Z" fill="#FFFFFF"/>'
    '<path d="M237.778 66L252 26H241.361L232 52.6841L222.645 26H212L226.222 66H237.778Z" fill="#FFFFFF"/>'
    '<path fill-rule="evenodd" clip-rule="evenodd" d="M208 6H198V28.6791C194.913 26.7289 191.922 26 188 26C176.954 26 168 34.9543 168 46C168 57.0457 176.954 66 188 66C191.643 66 195.058 65.0261 198 63.3243V66H208V46V6ZM198 46C198 40.4772 193.523 36 188 36C182.477 36 178 40.4772 178 46C178 51.5228 182.477 56 188 56C193.523 56 198 51.5228 198 46Z" fill="#FFFFFF"/>'
    '<path fill-rule="evenodd" clip-rule="evenodd" d="M154 28.5V26H164V66H154V63.5C151.055 65.2041 147.624 66 143.981 66C132.945 66 124 57.0464 124 46C124 34.9537 132.945 26 143.981 26C147.624 26 151.055 26.7959 154 28.5ZM144 56C149.523 56 154 51.5228 154 46C154 40.4772 149.523 36 144 36C138.477 36 134 40.4772 134 46C134 51.5228 138.477 56 144 56Z" fill="#FFFFFF"/>'
    '<path fill-rule="evenodd" clip-rule="evenodd" d="M80 6H70V46V66H80V63.3243C82.9417 65.0261 86.3571 66 90 66C101.046 66 110 57.0457 110 46C110 34.9543 101.046 26 90 26C86.3571 26 82.9417 26.9739 80 28.6757V6ZM80 46C80 51.5228 84.4772 56 90 56C95.5228 56 100 51.5228 100 46C100 40.4772 95.5228 36 90 36C84.4772 36 80 40.4772 80 46Z" fill="#FFFFFF"/>'
    '<path fill-rule="evenodd" clip-rule="evenodd" d="M65.9949 46.1337H65.9975V65.8967H65.9924C65.9939 65.9282 65.9954 65.9589 65.9964 65.9897C65.9971 66.0106 65.9975 66.0316 65.9975 66.053C65.9975 77.0686 57.0451 86 45.9987 86C34.9549 86 26 77.0706 26 66.053C26.0045 66.3535 26.0068 66.5068 26.007 66.5068C26.0071 66.5068 26.0048 66.3403 26 66H36C36 71.5092 40.4769 76 46 76C51.5232 76 55.9981 71.5622 55.9981 66.053C55.9981 66.0001 55.993 65.9471 55.993 65.8941H55.9981V63.3333C53.0569 65.0334 49.6442 66 46 66C34.9562 66 26 56.9648 26 45.9471C26 34.9293 34.9549 26 46.0013 26C49.6435 26 53.0585 26.9712 56 28.6682V26H66V45.9471V46C65.9996 46.0135 65.998 46.0269 65.9975 46.0403C65.9962 46.0712 65.9949 46.1022 65.9949 46.1337ZM35.9994 45.9471C35.9994 51.4559 40.4769 56 46 56C51.5232 56 56.0006 51.4559 55.9981 45.9471C55.9981 40.4381 51.4282 36 46 36C40.5718 36 35.9994 40.4381 35.9994 45.9471Z" fill="#FFFFFF"/>'
    '<path d="M19.4999 0H20.5C22.4885 9.79565 30.2045 17.5115 40 19.4999V20.5C30.2045 22.4885 22.4885 30.2045 20.5 40H19.4999C17.5115 30.2045 9.79565 22.4885 0 20.5V19.4999C9.79565 17.5115 17.5115 9.79565 19.4999 0Z" fill="#EA018B"/>'
    '</svg>'
)

STYLE = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{--magenta:#EA018B;--bright:#FF31BD;--purple:#504BE0;--ink:#12112C;--bg:#F8F9FA;--elev:#F1F3F5;
  --border:#DEE2E6;--muted:#495057;--subtle:#868E96;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Inter",system-ui,sans-serif;background:var(--bg);color:var(--ink);-webkit-font-smoothing:antialiased;}
.cover{background:conic-gradient(from 0deg at 50% 50%,var(--bright) 0%,var(--ink) 60%,var(--ink) 100%);
  color:#fff;padding:48px 56px 42px;}
.cover .logo{margin-bottom:26px;}
.cover h1{font-size:44px;font-weight:600;letter-spacing:-.02em;}
.cover .sub{font-size:17px;color:#ffd7f1;margin-top:8px;font-weight:500;}
.cover .meta{margin-top:18px;font-size:13px;color:#e7c9e6;display:flex;gap:18px;flex-wrap:wrap;}
.wrap{max-width:1600px;margin:0 auto;padding:32px 56px 48px;}
.krow{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px;}
.kpi{background:var(--elev);border-radius:18px;padding:16px 18px;}
.kpi .k{font-size:11.5px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;}
.kpi .v{font-size:28px;font-weight:600;letter-spacing:-.02em;margin-top:3px;font-variant-numeric:tabular-nums;}
.kpi .v.accent{color:var(--magenta);}.kpi .v.accent2{color:var(--purple);}
.kpi .d{font-size:12px;color:var(--subtle);margin-top:2px;}
.catbar{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:20px;font-size:12.5px;color:var(--muted);}
.catbar label.ctog{display:flex;align-items:center;gap:5px;cursor:pointer;}
.fdrop{position:relative;display:inline-block;}
.fdrop-btn{background:#fff;border:1px solid var(--border);border-radius:8px;padding:6px 12px;font-size:12.5px;
  cursor:pointer;color:var(--ink);font-family:inherit;display:inline-flex;align-items:center;gap:6px;}
.fdrop-btn:hover{background:var(--elev);}
.fdrop-btn .car{font-size:9px;color:var(--subtle);}
.fmenu{position:absolute;top:36px;left:0;background:#fff;border:1px solid var(--border);border-radius:10px;
  padding:8px;min-width:200px;max-height:260px;overflow:auto;display:none;z-index:30;
  box-shadow:0 10px 30px rgba(18,17,44,.15);}
.fmenu.open{display:block;}
.fmenu label{display:flex;align-items:center;gap:7px;padding:5px 6px;border-radius:6px;font-size:12.5px;
  cursor:pointer;white-space:nowrap;}
.fmenu label:hover{background:var(--elev);}
.fmenu .actrow{display:flex;gap:6px;padding:0 2px 8px;margin-bottom:6px;border-bottom:1px solid var(--border);}
.fmenu .actrow button{flex:1;font-size:11px;padding:5px;border-radius:6px;border:1px solid var(--border);
  background:var(--elev);cursor:pointer;font-family:inherit;color:var(--ink);}
.fmenu .actrow button:hover{background:#fff;}
.panel{background:var(--elev);border-radius:18px;padding:16px 18px;margin-bottom:20px;}
.ph{font-size:12.5px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:10px;
  display:flex;justify-content:space-between;align-items:baseline;}
.ph .tot{font-size:16px;font-weight:600;color:var(--ink);text-transform:none;letter-spacing:0;}
.tbl{width:100%;border-collapse:collapse;font-size:13px;background:#fff;border-radius:12px;overflow:hidden;}
.tbl th,.tbl td{text-align:left;padding:10px 12px;border-bottom:1px solid var(--border);vertical-align:top;}
.tbl th{font-size:10.5px;text-transform:uppercase;letter-spacing:.04em;color:var(--subtle);font-weight:600;background:#fff;}
.tbl td.n{text-align:right;font-variant-numeric:tabular-nums;}
.tbl tr:last-child td{border-bottom:none;}
.deal{font-weight:500;color:var(--ink);display:block;}
.sub{display:block;font-size:11.5px;color:var(--subtle);margin-top:1px;}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:10.5px;font-weight:600;white-space:nowrap;}
.b-neg{background:#fde7f4;color:#b00467;}.b-acc{background:#e9fbf1;color:#15803d;}.b-won{background:#dcfce7;color:#166534;}
.ctag{display:inline-block;font-size:9px;font-weight:700;letter-spacing:.03em;border-radius:4px;padding:2px 6px;
  color:#fff;margin-right:3px;}
.ct-cx{background:#EA018B;}.ct-ex{background:#504BE0;}.ct-d42{background:#0d9488;}.ct-ai{background:#d97706;}
.ct-pf{background:#0f766e;}.ct-ot{background:#868E96;}
.status-cell{max-width:220px;}
.status-cell .where{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px;cursor:help;}
.status-cell .next{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px;
  color:var(--subtle);font-size:11.5px;margin-top:2px;cursor:help;}
.empty{color:var(--subtle);font-size:13px;padding:6px 2px;}
.footer{margin-top:8px;font-size:12px;color:var(--subtle);}
</style>
"""

STAGE_BADGE = {"negotiation": "b-neg", "accepted proposal": "b-acc", "po invoiced": "b-won"}


def _ctags(categorias):
    return "".join(
        f'<span class="ctag {CAT_CLASS.get(c, "ct-ot")}">{esc(CAT_LABELS.get(c, c))}</span>' for c in categorias
    )


def fmt_total(v):
    return "$" + format(round(v), ",")


def _fila_html(f):
    cats_attr = " ".join(f["categorias"])
    stage_badge = STAGE_BADGE.get(f["stage"].strip().lower(), "b-acc")
    contacto = f["email"] or f["contacto"] or "-"
    where = f["where"].strip()
    next_step = f["next_step"].strip()
    where_html = f'<span class="where" title="{esc(where)}">{esc(where)}</span>' if where else "-"
    next_html = f'<span class="next" title="{esc(next_step)}">{esc(next_step)}</span>' if next_step else ""
    return f"""<tr data-cat="{esc(cats_attr)}" data-seg="{esc(f['segmento'])}" data-rep="{esc(f['owner'])}" data-country="{esc(f['pais'])}" data-amount="{f['amount']}">
      <td><span class="deal">{esc(f['deal'])}</span><span class="sub">{esc(f['organization'])}</span></td>
      <td>{esc(f['owner'])}</td>
      <td>{esc(f['pais'])}</td>
      <td><span class="badge {stage_badge}">{esc(f['stage'])}</span></td>
      <td>{f['ecd']:%b %d, %Y}</td>
      <td class="n">{f['amount_fmt']}</td>
      <td>{_ctags(f['categorias'])}</td>
      <td class="status-cell">{where_html}{next_html}</td>
      <td>{esc(contacto)}</td>
    </tr>"""


def _panel_forecast(clave, titulo, filas):
    total = sum(f["amount"] for f in filas)
    cuerpo = (
        "".join(_fila_html(f) for f in filas) if filas
        else '<tr><td colspan="9" class="empty">No deals in this period.</td></tr>'
    )
    return f"""
    <div class="panel">
      <div class="ph"><span id="ph-{clave}-count">{esc(titulo)} &middot; {len(filas)} deal(s)</span><span class="tot" id="ph-{clave}-total">{fmt_total(total)}</span></div>
      <table class="tbl" data-panel="{clave}">
        <tr><th>Deal / Organization</th><th>Rep</th><th>Country</th><th>Stage</th><th>ECD</th><th class="n">MRR</th><th>Category</th><th>Where are we? / Next Step</th><th>Primary Contact</th></tr>
        {cuerpo}
      </table>
    </div>"""


def _dropdown_filtro(id_base, label, css_class, opciones, seleccionadas_pred):
    """Filtro tipo Excel: un boton que abre un menu con checklist + Select
    All/None, en vez de una fila larga de checkboxes -- para listas que
    pueden tener muchos valores (paises, vendedores)."""
    items = "".join(
        f'<label><input type="checkbox" class="{css_class}" value="{esc(v)}" '
        f'{"checked" if seleccionadas_pred(v) else ""} onchange="applyF()"> {esc(v)}</label>'
        for v in opciones
    )
    return f"""<div class="fdrop">
      <button type="button" class="fdrop-btn" onclick="toggleMenu('{id_base}')">{esc(label)}
        <span class="car" id="{id_base}-count"></span> &#9662;</button>
      <div class="fmenu" id="{id_base}">
        <div class="actrow">
          <button type="button" onclick="selectAllIn('{css_class}', true)">All</button>
          <button type="button" onclick="selectAllIn('{css_class}', false)">None</button>
        </div>
        {items}
      </div>
    </div>"""


def formatear_periodo(meses_destino):
    return " & ".join(f"{MONTHS_EN[m - 1]} {a}" for a, m in sorted(meses_destino))


def build_html(gen_ns, upside, commit, meses_destino, incluir_servicios, generated_on,
               categorias_permitidas, segmentos_permitidos, reps_permitidos, paises_permitidos):
    date_str = generated_on.strftime("%B %d, %Y")
    mes_label = formatear_periodo(meses_destino)

    for f in upside + commit:
        f["amount_fmt"] = gen_ns.fmt_money(f["amount"])

    total_upside = sum(f["amount"] for f in upside)
    total_commit = sum(f["amount"] for f in commit)
    total_forecast = total_upside + total_commit

    panel_commit = _panel_forecast("commit", "Commit", commit)
    panel_upside = _panel_forecast("upside", "Upside", upside)

    # Solo se listan como filtro las categorias/segmentos/reps/paises que de
    # verdad tienen algun deal en este reporte -- una casilla que nunca
    # cambia nada (ej. EX cuando no hay ningun deal EX) solo confunde.
    todas_filas = upside + commit
    categorias_presentes = [c for c in CATEGORY_OPTIONS if any(c in f["categorias"] for f in todas_filas)]
    segmentos_presentes = [s for s in SEGMENT_OPTIONS if any(f["segmento"] == s for f in todas_filas)]
    reps_presentes = [r for r in gen_ns.TEAM_OWNERS if any(f["owner"] == r for f in todas_filas)]
    paises_presentes = sorted({f["pais"] for f in todas_filas})

    checks_cat = "".join(
        f'<label class="ctog"><input type="checkbox" class="fcat" value="{c}" '
        f'{"checked" if c in categorias_permitidas else ""} onchange="applyF()"> {esc(CAT_LABELS[c])}</label>'
        for c in categorias_presentes
    )
    checks_seg = "".join(
        f'<label class="ctog"><input type="checkbox" class="fseg" value="{esc(s)}" '
        f'{"checked" if s in segmentos_permitidos else ""} onchange="applyF()"> {esc(s)}</label>'
        for s in segmentos_presentes
    )
    dropdown_rep = _dropdown_filtro("menu-rep", "Rep", "frep", reps_presentes,
                                    lambda rep: rep in reps_permitidos)
    dropdown_country = _dropdown_filtro("menu-country", "Country", "fcountry", paises_presentes,
                                        lambda pais: paises_permitidos is None or pais in paises_permitidos)
    servicios_txt = "included" if incluir_servicios else "excluded"

    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>Forecast - {esc(mes_label)}</title>{STYLE}</head>
<body>
  <div class="cover">
    <div class="logo">{LOGO_SVG_WHITE}</div>
    <h1>Forecast</h1>
    <div class="sub">{esc(mes_label)} &middot; {esc(TEAM_LABEL)}</div>
    <div class="meta"><span>Generated on {date_str}</span><span>GB Services {servicios_txt}</span></div>
  </div>
  <div class="wrap">
    <div class="krow">
      <div class="kpi"><div class="k">Commit</div><div class="v accent" id="kpi-commit-total">{fmt_total(total_commit)}</div><div class="d" id="kpi-commit-count">{len(commit)} deal(s)</div></div>
      <div class="kpi"><div class="k">Upside</div><div class="v accent2" id="kpi-upside-total">{fmt_total(total_upside)}</div><div class="d" id="kpi-upside-count">{len(upside)} deal(s)</div></div>
      <div class="kpi"><div class="k">Total Forecast</div><div class="v" id="kpi-forecast-total">{fmt_total(total_forecast)}</div><div class="d" id="kpi-forecast-count">{len(commit) + len(upside)} deal(s)</div></div>
      <div class="kpi"><div class="k">GB Services</div><div class="v" style="font-size:20px">{servicios_txt.capitalize()}</div></div>
    </div>
    <div class="catbar">
      <span>Category:</span>
      {checks_cat}
    </div>
    <div class="catbar">
      <span>Segment:</span>
      {checks_seg}
    </div>
    <div class="catbar">
      {dropdown_rep}
      {dropdown_country}
    </div>
    {panel_commit}
    {panel_upside}
    <div class="footer">Generated automatically from DIQ and FW Cadence IQ ({date_str}). CX/EX category uses the same logic as the FW Cadence deck.</div>
  </div>
  <script>
  function money(n) {{ return "$" + Math.round(n).toLocaleString("en-US"); }}
  function applyF() {{
    var activeCat = Array.from(document.querySelectorAll('.fcat:checked')).map(function(i) {{ return i.value; }});
    var activeSeg = Array.from(document.querySelectorAll('.fseg:checked')).map(function(i) {{ return i.value; }});
    var activeRep = Array.from(document.querySelectorAll('.frep:checked')).map(function(i) {{ return i.value; }});
    var activeCountry = Array.from(document.querySelectorAll('.fcountry:checked')).map(function(i) {{ return i.value; }});
    var totals = {{}};
    ['commit', 'upside'].forEach(function(clave) {{
      var rows = document.querySelectorAll('table[data-panel="' + clave + '"] tr[data-cat]');
      var count = 0, sum = 0;
      rows.forEach(function(row) {{
        var cats = row.dataset.cat.split(' ');
        var catOk = cats.some(function(c) {{ return activeCat.indexOf(c) !== -1; }});
        var segOk = activeSeg.indexOf(row.dataset.seg) !== -1;
        var repOk = activeRep.indexOf(row.dataset.rep) !== -1;
        var countryOk = activeCountry.indexOf(row.dataset.country) !== -1;
        var visible = catOk && segOk && repOk && countryOk;
        row.style.display = visible ? '' : 'none';
        if (visible) {{ count++; sum += parseFloat(row.dataset.amount); }}
      }});
      totals[clave] = {{count: count, sum: sum}};
      var elCount = document.getElementById('ph-' + clave + '-count');
      if (elCount) elCount.textContent = (clave === 'commit' ? 'Commit' : 'Upside') + ' · ' + count + ' deal(s)';
      var elTotal = document.getElementById('ph-' + clave + '-total');
      if (elTotal) elTotal.textContent = money(sum);
    }});
    document.getElementById('kpi-commit-total').textContent = money(totals.commit.sum);
    document.getElementById('kpi-commit-count').textContent = totals.commit.count + ' deal(s)';
    document.getElementById('kpi-upside-total').textContent = money(totals.upside.sum);
    document.getElementById('kpi-upside-count').textContent = totals.upside.count + ' deal(s)';
    document.getElementById('kpi-forecast-total').textContent = money(totals.commit.sum + totals.upside.sum);
    document.getElementById('kpi-forecast-count').textContent = (totals.commit.count + totals.upside.count) + ' deal(s)';
    updateDropdownCount('menu-rep', 'frep');
    updateDropdownCount('menu-country', 'fcountry');
  }}
  function updateDropdownCount(menuId, cls) {{
    var all = document.querySelectorAll('.' + cls);
    var checked = document.querySelectorAll('.' + cls + ':checked');
    var el = document.getElementById(menuId + '-count');
    if (!el) return;
    el.textContent = checked.length === all.length ? '(All)' : '(' + checked.length + '/' + all.length + ')';
  }}
  function toggleMenu(id) {{
    document.querySelectorAll('.fmenu').forEach(function(m) {{ if (m.id !== id) m.classList.remove('open'); }});
    document.getElementById(id).classList.toggle('open');
  }}
  function selectAllIn(cls, on) {{
    document.querySelectorAll('.' + cls).forEach(function(cb) {{ cb.checked = on; }});
    applyF();
  }}
  document.addEventListener('click', function(e) {{
    if (!e.target.closest('.fdrop')) {{
      document.querySelectorAll('.fmenu').forEach(function(m) {{ m.classList.remove('open'); }});
    }}
  }});
  updateDropdownCount('menu-rep', 'frep');
  updateDropdownCount('menu-country', 'fcountry');
  </script>
</body></html>"""


def main():
    gen_ns = _load_gen_ns()

    parser = argparse.ArgumentParser(description="Forecast (Upside/Commit) for el responsable comercial's team.")
    parser.add_argument("--meses", default="actual",
                        help="Comma-separated months: 'actual', 'proximo' or 'YYYY-MM' "
                             "(e.g. actual,proximo or 2026-10,2026-12).")
    parser.add_argument("--desde", default=None,
                        help="Start of a month range ('actual', 'proximo' or YYYY-MM). "
                             "Use with --hasta; overrides --meses.")
    parser.add_argument("--hasta", default=None,
                        help="End of the month range, inclusive.")
    parser.add_argument("--excluir-servicios", action="store_true",
                        help="Exclude ProfessionalServices GB deals from the forecast.")
    parser.add_argument("--categorias", default=",".join(CATEGORY_OPTIONS),
                        help="Comma-separated categories to include: " + ",".join(CATEGORY_OPTIONS))
    parser.add_argument("--segmentos", default=",".join(SEGMENT_OPTIONS),
                        help="Comma-separated segments to include: " + ",".join(SEGMENT_OPTIONS))
    parser.add_argument("--reps", default=",".join(gen_ns.TEAM_OWNERS),
                        help="Comma-separated reps to include: " + ",".join(gen_ns.TEAM_OWNERS))
    parser.add_argument("--paises", default=None,
                        help="Comma-separated billing countries to include (default: all).")
    args = parser.parse_args()

    now = datetime.now()
    incluir_servicios = not args.excluir_servicios
    categorias_permitidas = {c.strip() for c in args.categorias.split(",") if c.strip()}
    segmentos_permitidos = {s.strip() for s in args.segmentos.split(",") if s.strip()}
    reps_permitidos = {r.strip() for r in args.reps.split(",") if r.strip()}
    paises_permitidos = (None if args.paises is None
                        else {p.strip() for p in args.paises.split(",") if p.strip()})
    try:
        if args.desde or args.hasta:
            # Un solo extremo basta: el que falte se iguala al otro, asi
            # "--desde 2026-11" significa exactamente ese mes.
            desde = args.desde or args.hasta
            hasta = args.hasta or args.desde
            meses_destino = rango_de_meses(desde, hasta, now)
            # El manifiesto guarda lo que se PIDIO, no solo lo resuelto (eso ya
            # va en "periodo"). Sin esta linea quedaba sin definir y el script
            # reventaba al escribir el manifiesto, despues de haber hecho todo
            # el trabajo (16-sep-2026).
            meses_elegidos = [desde] if desde == hasta else [desde, hasta]
        else:
            meses_elegidos = [m.strip() for m in args.meses.split(",") if m.strip()]
            meses_destino = meses_objetivo(meses_elegidos, now)
    except ValueError as exc:
        print(f"Error: {exc}")
        return
    if not meses_destino:
        print("No month selected (--meses / --desde / --hasta).")
        return

    diq_path = gen_ns.find_latest_files().get("diq")
    if not diq_path:
        print("Could not find the DIQ report in Downloads.")
        return
    fw_path = find_fw_cadence_deals(gen_ns)
    if not fw_path:
        print('Could not find "FW Cadence IQ_*.xlsx" in Downloads -- CX/EX categories and contact '
              "emails won't be available (everything will show as Payment Frequency or Other).")
        fw_map = {}
    else:
        fw_map = fw_info_por_deal_id(gen_ns, fw_path)

    print("Files used:")
    print(f"  diq: {diq_path.name}")
    print(f"  fw cadence iq: {fw_path.name if fw_path else '(not found)'}")

    diq_rows = gen_ns.read_xlsx_rows(diq_path)
    upside, commit = calcular_forecast(gen_ns, diq_rows, fw_map, meses_destino, incluir_servicios,
                                       categorias_permitidas, segmentos_permitidos, reps_permitidos,
                                       paises_permitidos)

    periodo_label = formatear_periodo(meses_destino)
    print(f"\nTarget period: {periodo_label}")
    print(f"GB Services: {'included' if incluir_servicios else 'excluded'}")
    print(f"Categories: {', '.join(sorted(categorias_permitidas))}")
    print(f"Segments: {', '.join(sorted(segmentos_permitidos))}")
    print(f"Reps: {', '.join(sorted(reps_permitidos))}")
    print(f"Countries: {', '.join(sorted(paises_permitidos)) if paises_permitidos is not None else 'all'}")
    print(f"Upside: {len(upside)} deal(s), {fmt_total(sum(f['amount'] for f in upside))}")
    print(f"Commit: {len(commit)} deal(s), {fmt_total(sum(f['amount'] for f in commit))}")

    html = build_html(gen_ns, upside, commit, meses_destino, incluir_servicios, now,
                      categorias_permitidas, segmentos_permitidos, reps_permitidos, paises_permitidos)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    nombre = f"Forecast - {periodo_label}_{now:%Y-%m-%d}.html"
    out_path = OUT_DIR / nombre
    out_path.write_text(html, encoding="utf-8")

    manifiesto = {
        "generado": now.isoformat(),
        "archivo": out_path.name,
        "meses": meses_elegidos,
        "periodo": sorted(f"{a}-{m:02d}" for a, m in meses_destino),
        "incluir_servicios": incluir_servicios,
        "categorias": sorted(categorias_permitidas),
        "segmentos": sorted(segmentos_permitidos),
        "reps": sorted(reps_permitidos),
        "paises": sorted(paises_permitidos) if paises_permitidos is not None else None,
        "upside_count": len(upside),
        "commit_count": len(commit),
    }
    MANIFEST_FILE.write_text(json.dumps(manifiesto, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\nReport: {out_path}")


if __name__ == "__main__":
    main()
