"""
Replica por API directa (sin pasar por el modulo Reports de vTiger, bloqueado
para esta integracion) los Excel de metricas Mid/SMB que hoy el responsable comercial descarga a
mano desde la UI de vTiger a su carpeta de Descargas.

Habla directo con webservice.php de vTiger (protocolo challenge/login/query
estandar), sin depender de ningun MCP ni de Claude en tiempo de ejecucion, asi
que este script se puede llamar solo desde un cron.

ESTADO DE LA REPLICA (validado 27-ago-2026 contra los 8 Excel reales bajados
ese dia a las 08:25-08:26, ver Descargas/):

  SOLIDOS (validados fila por fila, cuadran salvo drift normal de datos en
  vivo -- gente que se reasigna entre que corres el reporte real y este script):
    - orgiq: exacto. Accounts asignadas a los 4 del equipo.
    - churn (ChurnAssets): exacto. Assets con Operation='Cancellation',
      organizacion del equipo, y VT Change Date dentro de los ultimos 30 dias
      (ventana inferida por el hueco de fechas del archivo real del
      27-ago-2026: el corte cae entre 24-jul y 13-ago; 30 dias fue el valor
      mas razonable que reproduce EXACTO las 6 filas reales de ese dia -- si
      un dia la cuenta no cuadra con lo que el responsable comercial ve en vTiger, este es el
      primer numero a ajustar).
    - oa: exacto (mismo criterio que churn pero sin filtrar por Operation:
      todos los Assets de organizaciones del equipo).
    - far: exacto. Todo vtcmfarming asignado al equipo, cualquier stage/tipo.

  APROXIMADOS / SIN CONFIRMAR (no se pudo inferir el filtro exacto del reporte
  original en el tiempo disponible -- avisa cuando el conteo se desvie mucho
  de lo esperado en vez de asumir que esta bien):
    - diq: se usa "Deals asignados al equipo" (Potentials.assigned_user_id in
      team) sin filtrar por sales_stage. El 27-ago-2026 esto dio 2327 filas
      contra 2439 del reporte real (-4.6%). No se identifico el criterio que
      explica esa diferencia (no es por stage: se probaron variantes). Antes
      de usar esto para compensacion, el responsable comercial deberia confirmar si DIQ excluye
      algo especifico (ej. deals sin producto, o con browser un filtro de
      fecha) que no fue posible inferir de un solo snapshot.
    - created (Deals Created): mismo enfoque (Deals asignados al equipo) mas
      un filtro de fecha de creacion -- el rango exacto NO esta confirmado,
      se usa "createdtime dentro de los ultimos 90 dias" como aproximacion
      razonable, sin validar contra el archivo real (543 filas) por falta de
      tiempo. AJUSTAR / CONFIRMAR con el responsable comercial.
    - farmconver: se usa "Deals asignados al equipo, con Farming vinculado
      (cf_potentials_farming no vacio)". Sin validar contra el archivo real
      (654 filas) por falta de tiempo. AJUSTAR / CONFIRMAR con el responsable comercial.
    - farmnoprod (Farming sin Producto): INTENTADO Y DESCARTADO por ahora.
      El archivo real del 27-ago-2026 trae 1 sola fila; "Product vacio y
      stage abierto (no Completed/Completed Unattended)" da 113 candidatos
      en vTiger, asi que el reporte real tiene un filtro adicional que no se
      pudo identificar (no es fecha de creacion ni Farming Type obvio). Esta
      funcion queda escrita pero DESACTIVADA (genera un archivo vacio con el
      encabezado correcto) hasta que el responsable comercial confirme el criterio real o acepte
      la aproximacion amplia.

Uso:
    python3 "descargar_reportes_vtiger.py"                 # escribe en Descargas (default)
    python3 "descargar_reportes_vtiger.py" --output-dir /ruta/de/prueba
    python3 "descargar_reportes_vtiger.py" --solo orgiq,churn,oa,far
"""
import argparse
import hashlib
import json
import os
import urllib.parse
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

import load_env  # noqa: F401  (carga .env al importarse)
from paths import DOWNLOADS as DEFAULT_OUTPUT_DIR

VTIGER_BASE = "https://gbadvisors.od1.vtiger.com/webservice.php"

# Las credenciales estaban escritas en este archivo. Ahora salen de .env
# (VTIGER_USERNAME / VTIGER_ACCESS_KEY), que es el mismo par que usan el MCP
# de Vtiger y las rutinas -- una sola copia que rotar.
VTIGER_USER = os.environ.get("VTIGER_USERNAME", "")
VTIGER_ACCESS_KEY = os.environ.get("VTIGER_ACCESS_KEY", "")

if not VTIGER_USER or not VTIGER_ACCESS_KEY:
    raise SystemExit(
        "Faltan credenciales de vTiger. Define VTIGER_USERNAME y "
        "VTIGER_ACCESS_KEY en el archivo .env de la raiz del proyecto."
    )

# IDs internos de vTiger (modulo "Users", 19x...) del equipo Mid/SMB de el responsable comercial.
# El modulo Users esta bloqueado por API (ACCESS_DENIED, ver nota en
# Churn Accounts/.claude/skills/analiza-churn/SKILL.md); estos 4 se resolvieron
# el 27-ago-2026 anclando por una cuenta conocida de cada persona en Accounts
# y leyendo su assigned_user_id. Son identificadores internos permanentes de
# esas cuentas de usuario en vTiger -- no deberian cambiar salvo que se borre
# y recree el usuario.
TEAM = {
    "Vendedor 2": "19x222",
    "Vendedor 4": "19x260",
    "Vendedora 3": "19x425",
    "Vendedora 1": "19x431",
}
TEAM_IDS = list(TEAM.values())
ID_TO_NAME = {v: k for k, v in TEAM.items()}

CHURN_WINDOW_DAYS = 30  # ver nota "APROXIMADOS" arriba
CREATED_WINDOW_DAYS = 90  # sin confirmar, ver nota arriba

TZ_OFFSET_HOURS = -4  # vTiger API devuelve UTC; los reportes exportados muestran hora local (AST/GMT-4)


# ---------------- cliente vTiger (webservice.php nativo, sin MCP) ----------------

def login():
    r = urllib.request.urlopen(
        VTIGER_BASE + "?operation=getchallenge&username=" + urllib.parse.quote(VTIGER_USER),
        timeout=15,
    )
    challenge = json.loads(r.read())["result"]["token"]
    access_hash = hashlib.md5((challenge + VTIGER_ACCESS_KEY).encode()).hexdigest()
    data = urllib.parse.urlencode(
        {"operation": "login", "username": VTIGER_USER, "accessKey": access_hash}
    ).encode()
    r = urllib.request.urlopen(VTIGER_BASE, data=data, timeout=15)
    result = json.loads(r.read())
    if not result.get("success"):
        raise RuntimeError(f"Login fallo: {result}")
    return result["result"]["sessionName"]


def vquery(session, q):
    qs = urllib.parse.urlencode({"operation": "query", "sessionName": session, "query": q})
    r = urllib.request.urlopen(VTIGER_BASE + "?" + qs, timeout=30)
    result = json.loads(r.read())
    if not result.get("success"):
        raise RuntimeError(f"Query fallo: {result} -- query: {q}")
    return result["result"]


def vquery_all(session, select_from_where, batch=200, hard_cap=20000):
    """select_from_where: query VQL SIN 'LIMIT' ni ';' final, ej.
    "SELECT id, accountname FROM Accounts WHERE assigned_user_id IN ('19x222')" """
    out = []
    offset = 0
    while True:
        rows = vquery(session, f"{select_from_where} LIMIT {offset},{batch};")
        out.extend(rows)
        if len(rows) < batch or offset > hard_cap:
            break
        offset += batch
    return out


# ---------------- utilidades de formato ----------------

def utc_to_local_str(utc_str):
    """'2026-06-26 12:27:07' (UTC, tal como lo devuelve vTiger) ->
    '26-06-2026 08:27 AM' (local, mismo formato que exportan los reportes)."""
    if not utc_str:
        return None
    dt = datetime.strptime(utc_str, "%Y-%m-%d %H:%M:%S") + timedelta(hours=TZ_OFFSET_HOURS)
    return dt.strftime("%d-%m-%Y %I:%M %p").lstrip("0").replace(" 0", " ")


def utc_to_local_date(utc_str):
    """Para columnas que en el Excel real vienen como celda datetime (fecha,
    sin hora), ej. 'Assets VT Change Date', 'Deals Expected Close Date'."""
    if not utc_str:
        return None
    dt = datetime.strptime(utc_str, "%Y-%m-%d %H:%M:%S") + timedelta(hours=TZ_OFFSET_HOURS)
    return datetime(dt.year, dt.month, dt.day)


def num_or_none(v):
    if v in (None, ""):
        return None
    return float(v)


def name_or_id(user_id):
    return ID_TO_NAME.get(user_id, user_id)


# ---------------- Accounts del equipo (usado por varios reportes) ----------------

def fetch_team_accounts(session):
    fields = (
        "id, accountname, account_no, assigned_user_id, cf_1434, industry, "
        "createdtime, cf_accounts_nda, cf_accounts_engagementlevel, cf_944, "
        "bill_country, cf_accounts_gbslink, cf_accounts_ranking, "
        "cf_accounts_freshworkstier, cf_accounts_employeecount, website, "
        "last_contacted_on, annual_revenue, cf_accounts_mrrresult, "
        "accountstatus, cf_accounts_farminginactivereason"
    )
    ids = "','".join(TEAM_IDS)
    rows = vquery_all(session, f"SELECT {fields} FROM Accounts WHERE assigned_user_id IN ('{ids}')")
    return {r["id"]: r for r in rows}


def org_columns(acc):
    """Las 16 columnas 'Organizations *' compartidas por orgiq/oa/churn/far."""
    return {
        "Organizations Last Contacted On": utc_to_local_str(acc.get("last_contacted_on")),
        "Organizations Domain": acc.get("cf_1434") or None,
        "Organizations Industry": acc.get("industry") or None,
        "Organizations Created Time": utc_to_local_str(acc.get("createdtime")),
        "Organizations NDA": acc.get("cf_accounts_nda") or None,
        "Organizations Engagement Level": acc.get("cf_accounts_engagementlevel") or None,
        "Organizations Language": acc.get("cf_944") or None,
        "Organizations Billing Country": acc.get("bill_country") or None,
        "Organizations GBS Link": acc.get("cf_accounts_gbslink") or None,
        "Organizations Ranking": acc.get("cf_accounts_ranking") or None,
        "Organizations Freshworks Tier": acc.get("cf_accounts_freshworkstier") or None,
        "Organizations Employee Count": acc.get("cf_accounts_employeecount") or None,
        "Organizations Website": acc.get("website") or None,
        "Organizations Assigned To": name_or_id(acc.get("assigned_user_id")),
        "Organizations Organization ID": acc.get("account_no"),
        "Organizations Organization Name": acc.get("accountname"),
    }


# ---------------- 1. orgiq ----------------

def build_orgiq(session, acc_by_id):
    header = list(org_columns(next(iter(acc_by_id.values()))).keys()) if acc_by_id else []
    rows = [org_columns(acc) for acc in acc_by_id.values()]
    return header, rows


# ---------------- 3. churn (ChurnAssets) y 4. oa ----------------

def fetch_team_assets(session, acc_by_id):
    fields = "id, account, assetname, cf_assets_previousmrr, cf_assets_mbr, cf_assets_operation, cf_assets_previousmrrdate"
    all_assets = vquery_all(session, f"SELECT {fields} FROM Assets")
    return [a for a in all_assets if a.get("account") in acc_by_id]


def build_oa(acc_by_id, team_assets):
    header = [
        "Organizations Last Contacted On", "Assets Previous MRR", "Assets Current MRR",
        "Assets Asset Name", "Organizations Domain", "Organizations Industry",
        "Organizations Created Time", "Organizations NDA", "Organizations Engagement Level",
        "Organizations Language", "Organizations Billing Country", "Organizations GBS Link",
        "Organizations Ranking", "Organizations Freshworks Tier", "Organizations Employee Count",
        "Organizations Website", "Organizations Assigned To", "Organizations Organization ID",
        "Organizations Organization Name",
    ]
    rows = []
    for a in team_assets:
        acc = acc_by_id[a["account"]]
        oc = org_columns(acc)
        rows.append({
            "Organizations Last Contacted On": oc["Organizations Last Contacted On"],
            "Assets Previous MRR": num_or_none(a.get("cf_assets_previousmrr")),
            "Assets Current MRR": num_or_none(a.get("cf_assets_mbr")),
            "Assets Asset Name": a.get("assetname") or None,
            **{k: oc[k] for k in header if k.startswith("Organizations") and k != "Organizations Last Contacted On"},
        })
    return header, rows


def build_churn(acc_by_id, team_assets, today):
    header = [
        "Organizations Domain", "Assets Previous MRR", "Organizations MRR Value",
        "Organizations MRR Result", "Assets Current MRR", "Assets Asset Name",
        "Assets Operation", "Assets VT Change Date", "Organizations Industry",
        "Organizations Created Time", "Organizations Engagement Level", "Organizations Language",
        "Organizations Billing Country", "Organizations GBS Link", "Organizations Ranking",
        "Organizations Freshworks Tier", "Organizations Employee Count", "Organizations Website",
        "Organizations Assigned To", "Organizations Organization ID", "Organizations Organization Name",
    ]
    cutoff = today - timedelta(days=CHURN_WINDOW_DAYS)
    rows = []
    for a in team_assets:
        if (a.get("cf_assets_operation") or "").strip().lower() != "cancellation":
            continue
        change_date = utc_to_local_date(a.get("cf_assets_previousmrrdate") and a["cf_assets_previousmrrdate"] + " 00:00:00")
        if change_date is None or change_date.date() < cutoff:
            continue
        acc = acc_by_id[a["account"]]
        oc = org_columns(acc)
        rows.append({
            "Organizations Domain": oc["Organizations Domain"],
            "Assets Previous MRR": num_or_none(a.get("cf_assets_previousmrr")),
            "Organizations MRR Value": num_or_none(acc.get("annual_revenue")),
            "Organizations MRR Result": acc.get("cf_accounts_mrrresult") or None,
            "Assets Current MRR": num_or_none(a.get("cf_assets_mbr")),
            "Assets Asset Name": a.get("assetname") or None,
            "Assets Operation": "Cancellation",
            "Assets VT Change Date": change_date,
            **{k: oc[k] for k in ["Organizations Industry", "Organizations Created Time",
                                   "Organizations Engagement Level", "Organizations Language",
                                   "Organizations Billing Country", "Organizations GBS Link",
                                   "Organizations Ranking", "Organizations Freshworks Tier",
                                   "Organizations Employee Count", "Organizations Website",
                                   "Organizations Assigned To", "Organizations Organization ID",
                                   "Organizations Organization Name"]},
        })
    return header, rows


# ---------------- 5. far y 2. farmnoprod ----------------

def fetch_team_farming(session):
    fields = (
        "id, fld_vtcmfarmingname, cf_vtcmfarming_farmingstage, cf_vtcmfarming_farmingtype, "
        "assigned_user_id, vtcmfarmingnumber, cf_vtcmfarming_organization, "
        "cf_vtcmfarming_nextstep, cf_vtcmfarming_commfrequency, cf_vtcmfarming_product, modifiedtime"
    )
    ids = "','".join(TEAM_IDS)
    return vquery_all(session, f"SELECT {fields} FROM vtcmfarming WHERE assigned_user_id IN ('{ids}')")


def build_far(acc_by_id, team_farming):
    header = [
        "Farming Modified Time", "Organizations Account inactive reason", "Farming Next Step",
        "Organizations Account Status", "Farming Farming Name", "Farming Assigned To",
        "Farming Farming Number", "Farming Farming Type", "Farming Farming Stage",
        "Organizations Last Contacted On", "Organizations Domain", "Organizations Industry",
        "Organizations Created Time", "Organizations NDA", "Organizations Engagement Level",
        "Organizations Language", "Organizations Billing Country", "Organizations GBS Link",
        "Organizations Ranking", "Organizations Freshworks Tier", "Organizations Employee Count",
        "Organizations Website", "Organizations Assigned To", "Organizations Organization ID",
        "Organizations Organization Name",
    ]
    rows = []
    for f in team_farming:
        acc = acc_by_id.get(f.get("cf_vtcmfarming_organization"))
        oc = org_columns(acc) if acc else {k: None for k in header if k.startswith("Organizations")}
        rows.append({
            "Farming Modified Time": utc_to_local_str(f.get("modifiedtime")),
            "Organizations Account inactive reason": acc.get("cf_accounts_farminginactivereason") if acc else None,
            "Farming Next Step": f.get("cf_vtcmfarming_nextstep") or None,
            "Organizations Account Status": acc.get("accountstatus") if acc else None,
            "Farming Farming Name": f.get("fld_vtcmfarmingname"),
            "Farming Assigned To": name_or_id(f.get("assigned_user_id")),
            "Farming Farming Number": f.get("vtcmfarmingnumber") or None,
            "Farming Farming Type": f.get("cf_vtcmfarming_farmingtype") or None,
            "Farming Farming Stage": f.get("cf_vtcmfarming_farmingstage") or None,
            **{k: oc.get(k) for k in header if k.startswith("Organizations")},
        })
    return header, rows


def build_farmnoprod(team_farming, acc_by_id):
    """DESACTIVADO -- ver nota grande al inicio del archivo. Devuelve solo el
    encabezado (0 filas) hasta que se confirme el criterio real con el responsable comercial."""
    header = [
        "Farming Assigned To", "Farming Farming Name", "Products Product Name",
        "Farming Farming Stage", "Farming Organization", "Farming Comm Frequency",
    ]
    return header, []


# ---------------- 6/7/8. Potentials: farmconver / diq / created ----------------

def fetch_team_deals(session):
    fields = (
        "id, related_to, assigned_user_id, modifiedtime, nextstep, "
        "cf_potentials_companysegment, cf_potentials_wherearewe, "
        "cf_potentials_professionalservicetype, cf_potentials_assigneddate, "
        "cf_potentials_gbslink, closingdate, leadsource, cf_potentials_commfrequency, "
        "cf_948, sales_stage, amount, potentialname, potential_no, cf_potentials_farming, createdtime"
    )
    ids = "','".join(TEAM_IDS)
    return vquery_all(session, f"SELECT {fields} FROM Potentials WHERE assigned_user_id IN ('{ids}')")


def build_diq(acc_by_id, team_deals):
    header = [
        "Deals Modified Time", "Deals Next Step", "Deals Company Segment", "Deals Where are we?",
        "Deals Professional Service Type", "Deals Assigned Date", "Deals GBS Link",
        "Deals Assigned To", "Organizations GBS Link", "Deals Expected Close Date",
        "Deals Lead Source", "Deals Comm Frequency", "Deals New/Renew/ProServ",
        "Deals Sales Stage", "Deals Amount", "Deals Deal Name", "Deals Deal ID",
        "Products Product Name", "Deals Organization Name ",
    ]
    rows = []
    for d in team_deals:
        acc = acc_by_id.get(d.get("related_to"))
        rows.append({
            "Deals Modified Time": utc_to_local_str(d.get("modifiedtime")),
            "Deals Next Step": d.get("nextstep") or None,
            "Deals Company Segment": d.get("cf_potentials_companysegment") or None,
            "Deals Where are we?": d.get("cf_potentials_wherearewe") or None,
            "Deals Professional Service Type": d.get("cf_potentials_professionalservicetype") or None,
            "Deals Assigned Date": utc_to_local_str(d.get("cf_potentials_assigneddate")),
            "Deals GBS Link": d.get("cf_potentials_gbslink") or None,
            "Deals Assigned To": name_or_id(d.get("assigned_user_id")),
            "Organizations GBS Link": acc.get("cf_accounts_gbslink") if acc else None,
            "Deals Expected Close Date": utc_to_local_date(d.get("closingdate") and d["closingdate"] + " 00:00:00") if d.get("closingdate") else None,
            "Deals Lead Source": d.get("leadsource") or None,
            "Deals Comm Frequency": d.get("cf_potentials_commfrequency") or None,
            "Deals New/Renew/ProServ": d.get("cf_948") or None,
            "Deals Sales Stage": d.get("sales_stage") or None,
            "Deals Amount": num_or_none(d.get("amount")),
            "Deals Deal Name": d.get("potentialname"),
            "Deals Deal ID": d.get("potential_no") or None,
            "Products Product Name": None,  # requiere LineItem por deal, no incluido en esta pasada
            "Deals Organization Name ": acc.get("accountname") if acc else None,
        })
    return header, rows


def build_created(acc_by_id, team_deals, today):
    header = [
        "Deals Professional Service Type", "Deals Where are we?", "Deals Assigned Date",
        "Deals GBS Link", "Deals Assigned To", "Organizations GBS Link",
        "Deals Expected Close Date", "Deals Lead Source", "Deals Comm Frequency",
        "Deals New/Renew/ProServ", "Deals Sales Stage", "Deals Amount", "Deals Deal Name",
        "Deals Organization Name ", "Deals Created Time",
    ]
    cutoff = today - timedelta(days=CREATED_WINDOW_DAYS)
    rows = []
    for d in team_deals:
        created = utc_to_local_date(d.get("createdtime"))
        if created is None or created.date() < cutoff:
            continue
        acc = acc_by_id.get(d.get("related_to"))
        rows.append({
            "Deals Professional Service Type": d.get("cf_potentials_professionalservicetype") or None,
            "Deals Where are we?": d.get("cf_potentials_wherearewe") or None,
            "Deals Assigned Date": utc_to_local_str(d.get("cf_potentials_assigneddate")),
            "Deals GBS Link": d.get("cf_potentials_gbslink") or None,
            "Deals Assigned To": name_or_id(d.get("assigned_user_id")),
            "Organizations GBS Link": acc.get("cf_accounts_gbslink") if acc else None,
            "Deals Expected Close Date": utc_to_local_date(d.get("closingdate") and d["closingdate"] + " 00:00:00") if d.get("closingdate") else None,
            "Deals Lead Source": d.get("leadsource") or None,
            "Deals Comm Frequency": d.get("cf_potentials_commfrequency") or None,
            "Deals New/Renew/ProServ": d.get("cf_948") or None,
            "Deals Sales Stage": d.get("sales_stage") or None,
            "Deals Amount": num_or_none(d.get("amount")),
            "Deals Deal Name": d.get("potentialname"),
            "Deals Organization Name ": acc.get("accountname") if acc else None,
            "Deals Created Time": utc_to_local_str(d.get("createdtime")),
        })
    return header, rows


def build_farmconver(acc_by_id, team_deals):
    header = [
        "Deals New/Renew/ProServ", "Deals Professional Service Type", "Deals Organization Name ",
        "Deals Lead Source", "Organizations Assigned To", "Deals Farming",
        "Organizations Freshworks Tier", "Organizations Billing Country", "Deals Assigned To",
        "Deals Sales Stage", "Organizations Organization Name", "Deals Deal Name", "Deals Amount",
    ]
    rows = []
    for d in team_deals:
        if not d.get("cf_potentials_farming"):
            continue
        acc = acc_by_id.get(d.get("related_to"))
        rows.append({
            "Deals New/Renew/ProServ": d.get("cf_948") or None,
            "Deals Professional Service Type": d.get("cf_potentials_professionalservicetype") or None,
            "Deals Organization Name ": acc.get("accountname") if acc else None,
            "Deals Lead Source": d.get("leadsource") or None,
            "Organizations Assigned To": name_or_id(acc.get("assigned_user_id")) if acc else None,
            "Deals Farming": d.get("cf_potentials_farming"),
            "Organizations Freshworks Tier": acc.get("cf_accounts_freshworkstier") if acc else None,
            "Organizations Billing Country": acc.get("bill_country") if acc else None,
            "Deals Assigned To": name_or_id(d.get("assigned_user_id")),
            "Deals Sales Stage": d.get("sales_stage") or None,
            "Organizations Organization Name": acc.get("accountname") if acc else None,
            "Deals Deal Name": d.get("potentialname"),
            "Deals Amount": num_or_none(d.get("amount")),
        })
    return header, rows


# ---------------- escritura xlsx ----------------

def write_xlsx(path, header, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append([row.get(h) for h in header])
    for col in ws.columns:
        length = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max(length + 2, 10), 60)
    wb.save(path)


FILE_NAMES = {
    "orgiq": "MIDI Mid Market Metricas 2025 - ORGIQ_{date}.xlsx",
    "farmnoprod": "Metrics Farming sin Producto 2025_{date}.xlsx",
    "churn": "Mid Market Metricas 2025 - ChurnAssets_{date}.xlsx",
    "oa": "Mid Market Metricas 2025 - OA_{date}.xlsx",
    "far": "Mid Market Metricas 2025 - far_{date}.xlsx",
    "farmconver": "Mid Market Metricas 2025 - farmconver_{date}.xlsx",
    "diq": "Mid Market Metricas 2025 DIQ_{date}.xlsx",
    "created": "Mid Market Metricas 2025 Deals Created_{date}.xlsx",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR))
    parser.add_argument("--solo", default=None, help="lista separada por comas de categorias a generar")
    args = parser.parse_args()

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    categorias = args.solo.split(",") if args.solo else list(FILE_NAMES.keys())

    today = datetime.utcnow() + timedelta(hours=TZ_OFFSET_HOURS)
    stamp = today.strftime("%d-%m-%Y_%H%M")

    print("Autenticando contra vTiger...")
    session = login()

    print("Trayendo Accounts del equipo...")
    acc_by_id = fetch_team_accounts(session)
    print(f"  {len(acc_by_id)} cuentas")

    team_assets = None
    team_farming = None
    team_deals = None

    results = {}

    if "orgiq" in categorias:
        results["orgiq"] = build_orgiq(session, acc_by_id)

    if {"oa", "churn"} & set(categorias):
        print("Trayendo Assets del equipo...")
        team_assets = fetch_team_assets(session, acc_by_id)
        print(f"  {len(team_assets)} assets")
        if "oa" in categorias:
            results["oa"] = build_oa(acc_by_id, team_assets)
        if "churn" in categorias:
            results["churn"] = build_churn(acc_by_id, team_assets, today.date())

    if {"far", "farmnoprod"} & set(categorias):
        print("Trayendo Farming del equipo...")
        team_farming = fetch_team_farming(session)
        print(f"  {len(team_farming)} farmings")
        if "far" in categorias:
            results["far"] = build_far(acc_by_id, team_farming)
        if "farmnoprod" in categorias:
            results["farmnoprod"] = build_farmnoprod(team_farming, acc_by_id)

    if {"diq", "created", "farmconver"} & set(categorias):
        print("Trayendo Deals del equipo...")
        team_deals = fetch_team_deals(session)
        print(f"  {len(team_deals)} deals")
        if "diq" in categorias:
            results["diq"] = build_diq(acc_by_id, team_deals)
        if "created" in categorias:
            results["created"] = build_created(acc_by_id, team_deals, today.date())
        if "farmconver" in categorias:
            results["farmconver"] = build_farmconver(acc_by_id, team_deals)

    for cat, (header, rows) in results.items():
        fname = FILE_NAMES[cat].format(date=stamp)
        path = out_dir / fname
        write_xlsx(path, header, rows)
        print(f"{cat}: {len(rows)} filas -> {path}")


if __name__ == "__main__":
    main()
