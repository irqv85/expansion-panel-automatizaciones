"""
Reporte ejecutivo del departamento Mid/SMB, para compartir hacia arriba
(nivel C). Por defecto cubre el trimestre en curso hasta el momento de
ejecutarlo (QTD).

Lee los reportes que ya descargas (DIQ, ChurnAssets, OA) y produce un Excel
con el periodo, comparado contra el trimestre anterior completo y contra la
unica meta definida hasta ahora: la de suscripciones de ServiceNow FY26.

Metas (27-10-2026): Iran NO tiene metas definidas, salvo la de ServiceNow.
Hasta esa fecha el reporte traia las del OKR viejo de Freshworks Mid + Digital
(KR1 2.070.000, Payment Frequency 300.000, servicios 150.000, logos 20), que ya
no corresponden a Expansion Partners y se quitaron. Cuando haya mas metas, se
agregan junto a TARGET_SERVICENOW_*.

Reglas financieras (GB Advisors):
  - ARR: los servicios (ProfessionalServices / SSI add on) YA son valor anual,
    asi que ARR = monto. Todo lo demas es MRR mensual: ARR = monto x 12.
  - Categoria, segun el campo New/Renew/ProServ:
    "New" = NEW LOGO · "New-Add-On" y "New-Cross-Sell" = EXPANSION (caen sobre
    cuentas que ya son clientes) · "Payment Frequency" = su propia categoria,
    porque sumarlo a Expansion contaria el mismo ARR dos veces ·
    ProfessionalServices, "SSI add on" y "Renew SSI" = SERVICES (incluye
    renovaciones de servicio) · "Renewal" de licencia = RENEWAL, es del equipo
    de Renewals.
  - Ganado = Closed Won + PO Invoiced. Perdido = Closed Lost. Abierto = resto.
  - Suscripciones de ServiceNow (meta FY26): deals NEW LOGO o EXPANSION cuyo
    nombre dice ServiceNow o SN, en ARR (monto x 12) y SIN renovaciones. Asi
    cuadra con la diapositiva de ServiceNow: el Q3 de 2026 da 67.716 (el add-on
    SN FSM de Banreservas, 5.643 x 12).

Uso:
  python "Metricas Mes.py"                    Trimestre en curso al dia de hoy
  python "Metricas Mes.py" --trimestre 2026Q2  Un trimestre completo
  python "Metricas Mes.py" --mes 2026-07       Un mes especifico
"""
import argparse
import calendar
import importlib.util
import re
import sys
import traceback
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from paths import DOWNLOADS, METRICAS_DIR as OUTPUT_DIR

SCRIPT_DIR = Path(__file__).resolve().parent

# Quinta copia del equipo que habia escrita a mano en el proyecto. Con la lista
# equivocada el reporte no falla: sale vacio, que es peor (5-oct-2026).
from equipo import NOMBRE_PANEL, TEAM_OWNERS as EQUIPO

# ---- Metas ----
# Unica meta definida (diapositiva "FY26 target" de ServiceNow, 27-10-2026):
# suscripciones, 250.000 por trimestre y 1.000.000 en el ano, en ARR y sin
# renovaciones. Iran dijo que para lo logrado manda lo que hay en vTiger, no la
# diapositiva (el Q1 de ella decia 50.000 y vTiger tiene 25.299,96).
TARGET_SERVICENOW_Q = 250_000
TARGET_SERVICENOW_ANIO = 4 * TARGET_SERVICENOW_Q
PATRON_SERVICENOW = re.compile(r"service ?now|\bsn\b", re.IGNORECASE)

TIPOS_PROSERV = {"professionalservices gb", "professionalservices vendor", "ssi add on"}
TIPOS_RENEWAL = {"renewal"}
# El KR3 dice "incl. renewals": las renovaciones de servicio (Renew SSI) suman
# a Servicios. Las renovaciones de licencia son del equipo de Renewals.
TIPOS_SERVICIO_RENOVADO = {"renew ssi"}

# La categoria sale del campo New/Renew/ProServ, no del Lead Source: el valor
# exacto "New" es logo nuevo, mientras que Add-On, Cross-Sell y Payment
# Frequency caen sobre cuentas que ya son clientes, o sea expansion. Clasificar
# por Lead Source mandaba los "New" de origen farming a expansion y dejaba New
# Business en cero.
TIPO_NUEVO_LOGO = "new"
TIPOS_EXPANSION = {"new-add-on", "new-cross-sell"}
TIPO_PAYMENT_FREQ = "payment frequency"

ETAPAS_GANADAS = {"closed won", "po invoiced"}
ETAPAS_PERDIDAS = {"closed lost"}

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# ---- Estilo (paleta del dashboard comercial) ----
C_DARK, C_MID, C_SMB = "1F3864", "2E75B6", "1A4F72"
C_GREEN, C_YELLOW, C_RED, C_GREY = "70AD47", "FFC000", "C00000", "D9D9D9"
FONT_NAME = "Arial"
INK = "12112C"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '$#,##0;($#,##0);-'
MONEY2 = '$#,##0.00;($#,##0.00);-'
PCT = '0.0%'


def log(m):
    print(f"[{datetime.now():%H:%M:%S}] {m}", flush=True)


# ---------- datos ----------

def find_reporte(patron, desc, obligatorio=True):
    archivos = sorted(DOWNLOADS.glob(patron), key=lambda p: p.stat().st_mtime)
    if not archivos:
        if obligatorio:
            raise FileNotFoundError(f"Falta el reporte {desc} ({patron}) en {DOWNLOADS}.")
        return None
    return archivos[-1]


def leer(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    # Un export de vTiger sin resultados viene con CERO filas, ni siquiera la
    # cabecera, y next() reventaba con StopIteration. Como StopIteration no
    # lleva mensaje, el manejador de abajo imprimia "ERROR: " a secas y no
    # habia forma de saber que habia pasado. Caso real: el 1-oct-2026, primer
    # dia de Q4, el reporte de ChurnAssets salio vacio porque todavia no habia
    # churn en el trimestre, y el plan de compensacion fallaba sin explicar
    # nada (1-oct-2026).
    primera = next(it, None)
    if primera is None:
        wb.close()
        return [], []
    header = [str(h).strip() if h is not None else "" for h in primera]
    filas = [dict(zip(header, r)) for r in it if any(v not in (None, "") for v in r)]
    wb.close()
    return header, filas


def a_fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    t = str(v or "").strip()
    for f in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(t[:10], f).date()
        except ValueError:
            continue
    return None


def a_monto(v):
    if v in (None, ""):
        return 0.0
    try:
        return float(str(v).replace("$", "").replace(",", ""))
    except ValueError:
        return 0.0


def normalizar(filas):
    """Agrega tipo, categoria, ARR y estado a cada deal del equipo."""
    salida = []
    vistos = set()
    for r in filas:
        deal_id = str(r.get("Deals Deal ID") or "").strip()
        # DIQ trae "Products Product Name": un deal con varios productos
        # asociados sale en una fila por producto, con el mismo monto
        # repetido en cada una. Sin este filtro el ARR de ese deal se suma
        # dos o mas veces en los KR.
        if deal_id and deal_id in vistos:
            continue
        if deal_id:
            vistos.add(deal_id)

        vend = str(r.get("Deals Assigned To") or "").strip()
        tipo = str(r.get("Deals New/Renew/ProServ") or "").strip()
        tipo_l = tipo.lower()
        ls = str(r.get("Deals Lead Source") or "").strip()
        etapa = str(r.get("Deals Sales Stage") or "").strip()
        monto = a_monto(r.get("Deals Amount"))

        if tipo_l in TIPOS_PROSERV or tipo_l in TIPOS_SERVICIO_RENOVADO:
            arr, categoria = monto, "SERVICES"
        elif tipo_l in TIPOS_RENEWAL:
            arr, categoria = monto * 12, "RENEWAL"
        elif tipo_l == TIPO_NUEVO_LOGO:
            arr, categoria = monto * 12, "NEW LOGO"
        elif tipo_l in TIPOS_EXPANSION:
            arr, categoria = monto * 12, "EXPANSION"
        elif tipo_l == TIPO_PAYMENT_FREQ:
            arr, categoria = monto * 12, "PAYMENT FREQUENCY"
        else:
            arr, categoria = monto * 12, "SIN CLASIFICAR"

        etapa_l = etapa.lower()
        estado = ("GANADO" if etapa_l in ETAPAS_GANADAS
                  else "PERDIDO" if etapa_l in ETAPAS_PERDIDAS else "ABIERTO")

        salida.append({
            "deal_id": deal_id,
            "nombre": str(r.get("Deals Deal Name") or "").strip(),
            "org": str(r.get("Deals Organization Name") or "").strip(),
            "vendedor": vend, "tipo": tipo, "lead_source": ls, "etapa": etapa,
            "mrr": monto, "arr": arr, "categoria": categoria, "estado": estado,
            "ecd": a_fecha(r.get("Deals Expected Close Date")),
        })
    return salida


def rango_mes(anio, mes):
    return date(anio, mes, 1), date(anio, mes, calendar.monthrange(anio, mes)[1])


def mes_anterior(anio, mes):
    return (anio - 1, 12) if mes == 1 else (anio, mes - 1)


def trimestre_de(f):
    return (f.month - 1) // 3 + 1


def rango_trimestre(anio, q):
    ini = date(anio, 3 * (q - 1) + 1, 1)
    fin_mes = 3 * q
    return ini, date(anio, fin_mes, calendar.monthrange(anio, fin_mes)[1])


def trimestre_anterior(anio, q):
    return (anio - 1, 4) if q == 1 else (anio, q - 1)


def en_rango(deals, desde, hasta):
    return [d for d in deals if d["ecd"] and desde <= d["ecd"] <= hasta]


def suma(deals, campo="arr"):
    return sum(d[campo] for d in deals)


def _con_fuente(deals, fuente):
    """Copia cada deal agregandole de donde salio (equipo actual o legacy),
    para que la hoja de detalle anual pueda distinguirlos sin tocar los
    dicts originales."""
    return [dict(d, fuente=fuente) for d in deals]


def por_categoria(deals, cat):
    return [d for d in deals if d["categoria"] == cat]


def tasa(a, b):
    return (a / b) if b else None


def es_suscripcion_servicenow(d):
    """Suscripcion de ServiceNow para la meta FY26: logo nuevo o expansion, con
    ServiceNow o SN en el nombre del deal. Sin renovaciones, servicios ni
    payment frequency."""
    return d["categoria"] in ("NEW LOGO", "EXPANSION") and bool(PATRON_SERVICENOW.search(d["nombre"]))


def servicenow_por_trimestre(deals, anio):
    """{1: ARR, 2: ..., 3: ..., 4: ...} de suscripciones de ServiceNow ganadas
    en el ano, por trimestre de la fecha de cierre."""
    por_q = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    for d in deals:
        if es_suscripcion_servicenow(d) and d["ecd"] and d["ecd"].year == anio:
            por_q[trimestre_de(d["ecd"])] += d["arr"]
    return por_q


# ---------- estilo ----------

def celda(ws, fila, col, valor, *, bold=False, size=10, color=INK, fill=None,
          fmt=None, wrap=False, borde=True, align=None):
    c = ws.cell(row=fila, column=col, value=valor)
    c.font = Font(name=FONT_NAME, size=size, bold=bold, color=color)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        c.number_format = fmt
    if borde:
        c.border = BORDER
    c.alignment = Alignment(vertical="center" if not wrap else "top",
                            wrap_text=wrap, horizontal=align)
    return c


def banda(ws, fila, texto, ncols, color=C_DARK):
    for i in range(1, ncols + 1):
        celda(ws, fila, i, texto if i == 1 else None, bold=True, size=11,
              color="FFFFFF", fill=color, borde=False)
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)


def color_semaforo(pct):
    if pct is None:
        return C_GREY
    return C_GREEN if pct >= 0.8 else (C_YELLOW if pct >= 0.5 else C_RED)


# ---------- hojas ----------

def hoja_resumen(wb, ctx):
    ws = wb.create_sheet("Resumen Ejecutivo")
    ncols = 6
    m = ctx["periodo_label"]

    celda(ws, 1, 1, f"Mid/SMB - Reporte ejecutivo {m}", bold=True, size=16,
          color=C_DARK, borde=False)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    # El nombre sale de equipo.json: escrito a mano quedaba mintiendo en cuanto
    # el panel lo usaba otro equipo (5-oct-2026).
    _eq = NOMBRE_PANEL or "Equipo comercial"
    celda(ws, 2, 1, f"{_eq} - equipo de {len(EQUIPO)} vendedores | "
                    f"Generado {datetime.now():%d/%m/%Y %H:%M} | Fuente: {ctx['reporte'].name}",
          size=9, color="5B5F6B", borde=False)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)

    # --- A. Resultado del mes ---
    f = 4
    banda(ws, f, f"A. RESULTADO DEL PERIODO - {m.upper()}", ncols, C_SMB)
    f += 1
    for i, t in enumerate(["Concepto", "Monto CRM", "ARR ganado", "Deals",
                           f"ARR {ctx['periodo_prev_label']}", "Var. vs periodo anterior"], start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)
    f += 1

    fila_ini = f
    for etiqueta, cat in [("New Business (logos nuevos)", "NEW LOGO"),
                          ("Expansion (base instalada)", "EXPANSION"),
                          ("Servicios", "SERVICES"),
                          ("Renovaciones", "RENEWAL")]:
        act = por_categoria(ctx["mes_ganado"], cat)
        prev = por_categoria(ctx["prev_ganado"], cat)
        arr_act, arr_prev = suma(act), suma(prev)
        celda(ws, f, 1, etiqueta)
        celda(ws, f, 2, suma(act, "mrr"), fmt=MONEY2)
        celda(ws, f, 3, arr_act, fmt=MONEY)
        celda(ws, f, 4, len(act), align="center")
        celda(ws, f, 5, arr_prev, fmt=MONEY)
        var = celda(ws, f, 6, (arr_act - arr_prev), fmt=MONEY)
        var.font = Font(name=FONT_NAME, size=10, bold=True,
                        color=("1F5C2E" if arr_act >= arr_prev else C_RED))
        f += 1
    fila_fin = f - 1

    celda(ws, f, 1, "TOTAL GANADO", bold=True, fill="EFF3F8")
    for col, letra in ((2, "B"), (3, "C"), (4, "D"), (5, "E"), (6, "F")):
        c = celda(ws, f, col, f"=SUM({letra}{fila_ini}:{letra}{fila_fin})", bold=True,
                  fill="EFF3F8", fmt=(MONEY if col != 4 else None))
        if col == 2:
            c.number_format = MONEY2
    f += 2

    # --- B. Salud comercial del mes ---
    banda(ws, f, "B. SALUD COMERCIAL DEL PERIODO", ncols, C_SMB)
    f += 1
    ganados, perdidos = ctx["mes_ganado"], ctx["mes_perdido"]
    wr_q = tasa(len(ganados), len(ganados) + len(perdidos))
    wr_a = tasa(suma(ganados), suma(ganados) + suma(perdidos))
    ticket = tasa(suma(ganados), len(ganados))
    top = max(ganados, key=lambda d: d["arr"], default=None)
    concentracion = tasa(top["arr"], suma(ganados)) if top else None

    indicadores = [
        ("Win rate por cantidad", wr_q, PCT,
         f"{len(ganados)} ganados de {len(ganados) + len(perdidos)} cerrados"),
        ("Win rate por monto", wr_a, PCT,
         f"${suma(ganados):,.0f} ganados vs ${suma(perdidos):,.0f} perdidos"),
        ("Ticket promedio (ARR)", ticket, MONEY, f"{len(ganados)} deals ganados"),
        ("Concentracion del top deal", concentracion, PCT,
         (top["nombre"][:48] if top else "sin deals")),
        ("Pipeline abierto (ARR)", suma(ctx["abiertos"]), MONEY,
         f"{len(ctx['abiertos'])} deals activos, todas las fechas"),
    ]
    for etiqueta, valor, fmt, nota in indicadores:
        celda(ws, f, 1, etiqueta)
        c = celda(ws, f, 2, valor if valor is not None else "—", fmt=fmt, bold=True)
        if valor is None:
            c.number_format = "General"
        celda(ws, f, 3, nota, size=9, color="5B5F6B", wrap=True)
        ws.merge_cells(start_row=f, start_column=3, end_row=f, end_column=ncols)
        f += 1
    f += 1

    # --- C. Avance del periodo vs meta de ServiceNow ---
    # Unica meta definida (27-10-2026). Antes eran KR1 a KR3 del OKR viejo de
    # Freshworks Mid + Digital, que no corresponden a Expansion Partners.
    es_mes = ctx["es_mes"]
    banda(ws, f, "C. AVANCE DEL MES VS META DE SERVICENOW" if es_mes
          else "C. AVANCE DEL TRIMESTRE VS META DE SERVICENOW", ncols, C_SMB)
    f += 1
    for i, t in enumerate(["Metrica", "Ganado en el periodo",
                           "Meta del mes (1/3 del trimestre)" if es_mes else "Meta del trimestre",
                           "% de la meta", "Ritmo esperado a hoy", "Estado"], start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)
    f += 1

    avance = ctx["avance_meta"]
    meta_per = TARGET_SERVICENOW_Q / 3 if es_mes else TARGET_SERVICENOW_Q
    real = suma([d for d in ctx["mes_ganado"] if es_suscripcion_servicenow(d)])
    pct = tasa(real, meta_per)
    esperado = meta_per * avance
    celda(ws, f, 1, "ServiceNow · Suscripciones")
    celda(ws, f, 2, real, fmt=MONEY, bold=True)
    c = celda(ws, f, 3, meta_per, fmt=MONEY)
    c.comment = Comment("Meta FY26 de ServiceNow: 250.000 por trimestre, 1.000.000 en el ano. "
                        "Logrado = logos nuevos y expansion con ServiceNow o SN en el nombre, "
                        "en ARR, sin renovaciones.", "Metricas Mes")
    cp = celda(ws, f, 4, f"=B{f}/C{f}", fmt=PCT, bold=True)
    cp.fill = PatternFill("solid", fgColor=color_semaforo(pct))
    cp.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    c2 = celda(ws, f, 5, esperado, fmt=MONEY, color="5B5F6B")
    c2.comment = Comment(f"Meta x {avance:.0%} del periodo transcurrido "
                         f"({ctx['dias_corridos_meta']} de {ctx['dias_meta']} dias).", "Metricas Mes")
    if avance >= 1:
        # Periodo ya cerrado: no hay "ritmo", se cumplio o no.
        estado_c = "Meta cumplida" if real >= meta_per else "Meta no cumplida"
    else:
        estado_c = "Sobre ritmo" if real >= esperado else "Bajo el ritmo esperado"
    celda(ws, f, 6, estado_c, size=9, color=("1F5C2E" if real >= esperado else C_RED), wrap=True)
    f += 1
    celda(ws, f, 1, f"Transcurrido {avance:.0%} del {'mes' if es_mes else 'trimestre'} "
                    f"({ctx['dias_corridos_meta']} de {ctx['dias_meta']} dias). Solo ServiceNow "
                    f"tiene meta definida; los demas fabricantes y los servicios todavia no.",
          size=9, color="5B5F6B", wrap=True)
    ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=ncols)
    f += 2

    # --- D. Avance anual de ServiceNow, por trimestre ---
    # Misma forma que la diapositiva "FY26 target: total and by quarter" de
    # ServiceNow. Lo logrado sale de vTiger (decision de Iran, 27-10-2026).
    anio_meta = ctx["anio_ref"]
    banda(ws, f, f"D. SERVICENOW FY{str(anio_meta)[-2:]}: META Y LOGRADO POR TRIMESTRE (suscripciones, ARR)",
          ncols, C_SMB)
    f += 1
    for i, t in enumerate(["Trimestre", "Meta", "Logrado (ARR)", "% de la meta",
                           "Ritmo esperado a hoy", "Estado"], start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)
    f += 1

    por_q = servicenow_por_trimestre(ctx["ytd_ganado"], anio_meta)
    ref = ctx["fecha_ref"]
    if ref.year > anio_meta:
        q_cerrados, q_curso = 4, None
    elif ref.year < anio_meta:
        q_cerrados, q_curso = 0, None
    else:
        q_ref = trimestre_de(ref)
        if ref >= rango_trimestre(anio_meta, q_ref)[1]:
            q_cerrados, q_curso = q_ref, None      # el trimestre ya termino
        else:
            q_cerrados, q_curso = q_ref - 1, q_ref
    for q in (1, 2, 3, 4):
        real_q = por_q[q]
        pct = tasa(real_q, TARGET_SERVICENOW_Q)
        celda(ws, f, 1, f"Q{q} {anio_meta}")
        celda(ws, f, 2, TARGET_SERVICENOW_Q, fmt=MONEY)
        celda(ws, f, 3, real_q, fmt=MONEY, bold=True)
        cp = celda(ws, f, 4, f"=C{f}/B{f}", fmt=PCT, bold=True)
        if q <= q_cerrados or q == q_curso:
            cp.fill = PatternFill("solid", fgColor=color_semaforo(pct))
            cp.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        if q == q_curso:
            esperado_q = TARGET_SERVICENOW_Q * ctx["avance_trimestre_actual"]
            celda(ws, f, 5, esperado_q, fmt=MONEY, color="5B5F6B")
            celda(ws, f, 6, "En curso: sobre ritmo" if real_q >= esperado_q else "En curso: bajo el ritmo esperado",
                  size=9, color=("1F5C2E" if real_q >= esperado_q else C_RED), wrap=True)
        else:
            celda(ws, f, 5, None)
            if q <= q_cerrados:
                cumplida = real_q >= TARGET_SERVICENOW_Q
                celda(ws, f, 6, "Meta cumplida" if cumplida else "Meta no cumplida", size=9,
                      color=("1F5C2E" if cumplida else C_RED), wrap=True)
            else:
                celda(ws, f, 6, "Por llegar", size=9, color="5B5F6B", wrap=True)
        f += 1
    total_real = sum(por_q.values())
    ritmo = ctx["meses_transcurridos"] / 12
    celda(ws, f, 1, f"FY{str(anio_meta)[-2:]} total", bold=True, fill="EFF3F8")
    celda(ws, f, 2, TARGET_SERVICENOW_ANIO, fmt=MONEY, bold=True, fill="EFF3F8")
    celda(ws, f, 3, total_real, fmt=MONEY, bold=True, fill="EFF3F8")
    cp = celda(ws, f, 4, f"=C{f}/B{f}", fmt=PCT, bold=True)
    cp.fill = PatternFill("solid", fgColor=color_semaforo(tasa(total_real, TARGET_SERVICENOW_ANIO)))
    cp.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    celda(ws, f, 5, TARGET_SERVICENOW_ANIO * ritmo, fmt=MONEY, color="5B5F6B", fill="EFF3F8")
    celda(ws, f, 6, "Sobre ritmo" if total_real >= TARGET_SERVICENOW_ANIO * ritmo else "Bajo el ritmo esperado",
          size=9, color=("1F5C2E" if total_real >= TARGET_SERVICENOW_ANIO * ritmo else C_RED), wrap=True, fill="EFF3F8")
    f += 1

    pipe_sn = suma([d for d in ctx["abiertos"] if es_suscripcion_servicenow(d)])
    n_sn = len([d for d in ctx["abiertos"] if es_suscripcion_servicenow(d)])
    celda(ws, f, 1, "Pipeline de ServiceNow abierto")
    celda(ws, f, 2, pipe_sn, fmt=MONEY, bold=True)
    celda(ws, f, 3, f"{n_sn} deals", size=9, color="5B5F6B")
    celda(ws, f, 4, None)
    celda(ws, f, 5, None)
    celda(ws, f, 6, f"Falta para la meta anual: ${max(TARGET_SERVICENOW_ANIO - total_real, 0):,.0f}",
          size=9, color="5B5F6B", wrap=True)
    f += 2

    # --- D. Churn ---
    banda(ws, f, "E. CHURN DEL PERIODO", ncols, C_SMB)
    f += 1
    ch = ctx["churn"]
    ch_si = [c for c in ch if c["corresponde"] == "Si"]
    ch_rev = [c for c in ch if c["corresponde"] == "Revisar"]
    resumen_churn = [
        ("Assets revisados", len(ch), None),
        ("Churn que corresponde al equipo", sum(c["perdido"] for c in ch_si), MONEY2),
        ("Casos por revisar", len(ch_rev), None),
    ]
    for etiqueta, valor, fmt in resumen_churn:
        celda(ws, f, 1, etiqueta)
        celda(ws, f, 2, valor, fmt=fmt, bold=True)
        f += 1
    celda(ws, f, 1, "Regla: el churn aplica solo si no queda ningun asset activo en la "
                    "organizacion. Detalle en la hoja Churn.", size=9, color="5B5F6B", wrap=True)
    ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=ncols)
    f += 2

    # --- E. Lectura ejecutiva ---
    banda(ws, f, "F. LECTURA EJECUTIVA", ncols, C_DARK)
    f += 1
    for texto in ctx["insights"]:
        c = celda(ws, f, 1, "• " + texto, wrap=True, borde=False)
        ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=ncols)
        ws.row_dimensions[f].height = 30
        f += 1

    for col, ancho in zip("ABCDEF", [38, 17, 17, 13, 17, 30]):
        ws.column_dimensions[col].width = ancho
    ws.sheet_view.showGridLines = False
    return ws


def hoja_vendedores(wb, ctx):
    ws = wb.create_sheet("Por Vendedor")
    celda(ws, 1, 1, f"Desempeno por vendedor - {ctx['periodo_label']}", bold=True, size=14,
          color=C_DARK, borde=False)
    encabezados = ["Vendedor", "New Business ARR", "Expansion ARR", "Servicios ARR",
                   "Total ARR ganado", "Monto CRM", "Ganados", "Perdidos",
                   "Win rate cant.", "Win rate monto", "Aporte % del equipo"]
    for i, t in enumerate(encabezados, start=1):
        celda(ws, 3, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)

    total_equipo = suma(ctx["mes_ganado"]) or 1
    f = 4
    fila_ini = f
    for vend in EQUIPO:
        g = [d for d in ctx["mes_ganado"] if d["vendedor"] == vend]
        p = [d for d in ctx["mes_perdido"] if d["vendedor"] == vend]
        celda(ws, f, 1, vend)
        celda(ws, f, 2, suma(por_categoria(g, "NEW LOGO")), fmt=MONEY)
        celda(ws, f, 3, suma(por_categoria(g, "EXPANSION")), fmt=MONEY)
        celda(ws, f, 4, suma(por_categoria(g, "SERVICES")), fmt=MONEY)
        celda(ws, f, 5, f"=SUM(B{f}:D{f})", fmt=MONEY, bold=True)
        celda(ws, f, 6, suma(g, "mrr"), fmt=MONEY2)
        celda(ws, f, 7, len(g), align="center")
        celda(ws, f, 8, len(p), align="center")
        wq = tasa(len(g), len(g) + len(p))
        wa = tasa(suma(g), suma(g) + suma(p))
        for col, val in ((9, wq), (10, wa)):
            c = celda(ws, f, col, val if val is not None else "—", fmt=PCT)
            if val is None:
                c.number_format = "General"
        celda(ws, f, 11, suma(g) / total_equipo, fmt=PCT)
        f += 1
    fila_fin = f - 1

    celda(ws, f, 1, "TOTAL EQUIPO", bold=True, fill="EFF3F8")
    for col in range(2, 12):
        letra = get_column_letter(col)
        if col in (9, 10):
            celda(ws, f, col, None, fill="EFF3F8")
            continue
        c = celda(ws, f, col, f"=SUM({letra}{fila_ini}:{letra}{fila_fin})", bold=True,
                  fill="EFF3F8", fmt=(MONEY if col in (2, 3, 4, 5) else None))
        if col == 6:
            c.number_format = MONEY2
        if col == 11:
            c.number_format = PCT

    for i, ancho in enumerate([22, 17, 16, 15, 17, 15, 10, 10, 13, 13, 15], start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = ws.cell(row=4, column=1)
    ws.sheet_view.showGridLines = False
    return ws


def hoja_pipeline(wb, ctx):
    ws = wb.create_sheet("Pipeline")
    celda(ws, 1, 1, "Pipeline abierto del equipo", bold=True, size=14, color=C_DARK, borde=False)
    celda(ws, 2, 1, "Todos los deals activos (no ganados ni perdidos), sin filtro de fecha.",
          size=9, color="5B5F6B", borde=False)

    abiertos = ctx["abiertos"]
    orden = ["Negotiation", "PO Invoiced", "Accepted Proposal", "Quote or Proposal",
             "Demo or POC", "Why Analysis", "Qualification"]
    for i, t in enumerate(["Etapa", "Deals", "ARR", "% del pipeline"], start=1):
        celda(ws, 4, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID)
    total = suma(abiertos) or 1
    etapas = sorted({d["etapa"] for d in abiertos},
                    key=lambda e: orden.index(e) if e in orden else 99)
    f = 5
    ini = f
    for etapa in etapas:
        grupo = [d for d in abiertos if d["etapa"] == etapa]
        celda(ws, f, 1, etapa)
        celda(ws, f, 2, len(grupo), align="center")
        celda(ws, f, 3, suma(grupo), fmt=MONEY)
        celda(ws, f, 4, suma(grupo) / total, fmt=PCT)
        f += 1
    celda(ws, f, 1, "TOTAL", bold=True, fill="EFF3F8")
    celda(ws, f, 2, f"=SUM(B{ini}:B{f-1})", bold=True, fill="EFF3F8", align="center")
    celda(ws, f, 3, f"=SUM(C{ini}:C{f-1})", bold=True, fill="EFF3F8", fmt=MONEY)
    celda(ws, f, 4, f"=SUM(D{ini}:D{f-1})", bold=True, fill="EFF3F8", fmt=PCT)

    f += 2
    celda(ws, f, 1, "Top 15 deals abiertos por ARR", bold=True, size=12, color=C_DARK, borde=False)
    f += 1
    for i, t in enumerate(["Deal ID", "Oportunidad", "Organizacion", "Vendedor",
                           "Etapa", "Categoria", "ECD", "ARR"], start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)
    f += 1
    for d in sorted(abiertos, key=lambda x: -x["arr"])[:15]:
        celda(ws, f, 1, d["deal_id"])
        celda(ws, f, 2, d["nombre"], wrap=True)
        celda(ws, f, 3, d["org"])
        celda(ws, f, 4, d["vendedor"])
        celda(ws, f, 5, d["etapa"])
        celda(ws, f, 6, d["categoria"])
        celda(ws, f, 7, f"{d['ecd']:%d/%m/%Y}" if d["ecd"] else "")
        celda(ws, f, 8, d["arr"], fmt=MONEY)
        f += 1

    for i, ancho in enumerate([20, 44, 28, 20, 18, 16, 12, 15], start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.sheet_view.showGridLines = False
    return ws


def hoja_deals(wb, ctx):
    ws = wb.create_sheet("Deals del Mes")
    celda(ws, 1, 1, f"Deals cerrados en {ctx['periodo_label']} (detalle auditable)", bold=True,
          size=14, color=C_DARK, borde=False)
    cols = ["Deal ID", "Oportunidad", "Organizacion", "Vendedor", "Etapa", "Tipo",
            "Lead Source", "Categoria", "ECD", "Monto CRM", "ARR"]
    for i, t in enumerate(cols, start=1):
        celda(ws, 3, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)
    f = 4
    for d in sorted(ctx["mes_ganado"] + ctx["mes_perdido"],
                    key=lambda x: (x["estado"] != "GANADO", -x["arr"])):
        fill = "E4F3EA" if d["estado"] == "GANADO" else None
        for i, v in enumerate([d["deal_id"], d["nombre"], d["org"], d["vendedor"], d["etapa"],
                               d["tipo"], d["lead_source"], d["categoria"],
                               f"{d['ecd']:%d/%m/%Y}" if d["ecd"] else "",
                               d["mrr"], d["arr"]], start=1):
            c = celda(ws, f, i, v, fill=fill, wrap=(i == 2))
            if i == 10:
                c.number_format = MONEY2
            if i == 11:
                c.number_format = MONEY
        f += 1
    for i, ancho in enumerate([20, 44, 28, 20, 16, 24, 30, 15, 12, 14, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = ws.cell(row=4, column=1)
    ws.sheet_view.showGridLines = False
    return ws


# ---------- farming ----------

# Sofia Garcia lleva el farming de todo GB y va aparte del equipo: sirve de
# referencia, no suma al equipo (decision de Iran, 27-10-2026).
SOFIA = "Sofia Garcia"
FABRICANTES_FOCO = [
    ("Halo", re.compile(r"\bhalo", re.IGNORECASE)),
    ("Humand", re.compile(r"humand", re.IGNORECASE)),
    ("Atera", re.compile(r"\batera\b", re.IGNORECASE)),
    ("NinjaOne", re.compile(r"ninja", re.IGNORECASE)),
    ("ServiceNow", PATRON_SERVICENOW),
]
FARMING_ACTIVO_NO = {"completed", "farming unattended", "farming unnatended",
                     "completed unattended", "completed unnatended"}
TIPOS_FARMING = ["Cross-Sell", "Upsell", "Starting Farming"]


def fabricante_farming(producto, nombre):
    """Fabricante de un farming. Farming Product viene vacio en ~1 de cada 4 y
    el nombre suele decirlo ("2026 - HaloITSM - Banco X"); si ni asi se sabe,
    queda "Sin producto" en vez de adivinar."""
    for texto in (producto, nombre):
        if not texto:
            continue
        for etiqueta, patron in FABRICANTES_FOCO:
            if patron.search(str(texto)):
                return etiqueta
    # Los servicios son servicios y ya (Iran, 27-10-2026): no se mezclan con
    # licencias ni cuentan como "otro fabricante".
    if re.search(r"services|ssi|sip|booster", f"{producto} {nombre}", re.IGNORECASE):
        return "Servicios"
    return "Fuera de foco" if producto else "Sin producto"


def cargar_farmings(reporte):
    """Un farming por (nombre, organizacion): Clauco Org lo repite por cada
    asset de la organizacion."""
    _, filas = leer(reporte)
    vistos, out = set(), []
    for r in filas:
        nombre = r.get("Farming Farming Name")
        org = r.get("Organizations Organization Name") or ""
        if not nombre or (nombre, org) in vistos:
            continue
        vistos.add((nombre, org))
        etapa = str(r.get("Farming Farming Stage") or "").strip()
        producto = str(r.get("Farming Product") or "").strip()
        out.append({
            "farming": nombre, "org": org,
            "vendedor": r.get("Farming Assigned To") or "",
            "tipo": str(r.get("Farming Farming Type") or "").strip() or "Sin tipo",
            "etapa": etapa or "Sin etapa",
            "activo": etapa.lower().rstrip(".").strip() not in FARMING_ACTIVO_NO,
            "producto": producto, "fabricante": fabricante_farming(producto, nombre),
            "empleados": str(r.get("Organizations Employee Count") or "").strip(),
            "pais": r.get("Organizations Billing Country") or "",
            "creado": a_fecha(r.get("Farming Created Time")),
            "modificado": a_fecha(r.get("Farming Modified Time")),
        })
    return out


def hoja_farming(wb, ctx):
    ws = wb.create_sheet("Farming")
    celda(ws, 1, 1, "Farming del equipo", bold=True, size=14, color=C_DARK, borde=False)
    farmings = ctx.get("farmings")
    if farmings is None:
        celda(ws, 2, 1, "Falta el reporte Clauco Org en Descargas (se baja en la corrida de la "
              "manana de Clauco). Bajalo y vuelve a generar.", size=10, color=C_RED, borde=False)
        ws.sheet_view.showGridLines = False
        return ws
    desde, hasta = ctx["desde"], ctx["hasta"]
    celda(ws, 2, 1, f"Fuente: {ctx['farmings_fuente']}. \"Creados\" es del periodo {ctx['periodo_label']}; "
          "\"Activos\" es la foto de hoy (todo farming que no esta Completed). Sofia Garcia va aparte, "
          "como referencia.", size=9, color="5B5F6B", borde=False)

    def activos(lista):
        return [x for x in lista if x["activo"]]

    def creados(lista):
        return [x for x in lista if x["creado"] and desde <= x["creado"] <= hasta]

    def de(vendedor):
        return [x for x in farmings if x["vendedor"] == vendedor]

    equipo = [x for x in farmings if x["vendedor"] in EQUIPO]
    sofia = de(SOFIA)

    # --- A. Por vendedor ---
    banda(ws, 4, "A. Por vendedor", 7, C_SMB)
    enc = ["Vendedor", "Creados en el periodo", "Activos hoy"] + TIPOS_FARMING + ["Otros tipos"]
    for i, t in enumerate(enc, start=1):
        celda(ws, 5, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True, align="center" if i > 1 else None)
    f = 6
    filas_a = [(v, de(v)) for v in EQUIPO] + [("Equipo", equipo), ("Sofia Garcia (aparte)", sofia)]
    for nombre, lista in filas_a:
        total = nombre == "Equipo"
        fill = "EFF3F8" if total else ("F5F5F5" if nombre.startswith("Sofia") else None)
        act = activos(lista)
        celda(ws, f, 1, nombre, bold=total, fill=fill)
        celda(ws, f, 2, len(creados(lista)), bold=total, fill=fill, align="center")
        celda(ws, f, 3, len(act), bold=total, fill=fill, align="center")
        for j, tipo in enumerate(TIPOS_FARMING, start=4):
            celda(ws, f, j, sum(1 for x in act if x["tipo"] == tipo), bold=total, fill=fill, align="center")
        celda(ws, f, 7, sum(1 for x in act if x["tipo"] not in TIPOS_FARMING), bold=total, fill=fill, align="center")
        f += 1
    fin_a = f - 1
    n_vend = len(EQUIPO)

    # --- B. Por fabricante ---
    f += 1
    banda(ws, f, "B. Por fabricante (cross-sell es lo que se quiere vender)", 7, C_SMB)
    f += 1
    for i, t in enumerate(["Fabricante", "Equipo: activos", "Equipo: cross-sell activos",
                           "Equipo: creados en el periodo", "Sofia: activos",
                           "Sofia: cross-sell activos", "Sofia: creados en el periodo"], start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True, align="center" if i > 1 else None)
    f += 1
    fabricantes = [n for n, _ in FABRICANTES_FOCO] + ["Servicios", "Fuera de foco", "Sin producto"]
    for fab in fabricantes:
        e = [x for x in equipo if x["fabricante"] == fab]
        s = [x for x in sofia if x["fabricante"] == fab]
        fill = "F5F5F5" if fab in ("Servicios", "Fuera de foco", "Sin producto") else None
        celda(ws, f, 1, fab, fill=fill)
        for j, v in enumerate([len(activos(e)),
                               sum(1 for x in activos(e) if x["tipo"] == "Cross-Sell"),
                               len(creados(e)), len(activos(s)),
                               sum(1 for x in activos(s) if x["tipo"] == "Cross-Sell"),
                               len(creados(s))], start=2):
            celda(ws, f, j, v, fill=fill, align="center")
        f += 1
    celda(ws, f, 1, "Atera aun no tiene farmings. Si el farming tiene Halo cuenta como Halo aunque tenga Freshworks. "
          "\"Fuera de foco\": Freshworks, monday, Intercom, Vtiger y similares. \"Sin producto\": no se adivina.", size=9, color="5B5F6B", borde=False)

    # --- D. Por tamano de empresa ---
    f += 2
    banda(ws, f, "D. Por tamano de la empresa (empleados de la organizacion)", 7, C_SMB)
    f += 1
    for i, t in enumerate(["Empleados", "Equipo: activos", "Equipo: cross-sell activos",
                           "Equipo: de foco activos", "Sofia: activos",
                           "Sofia: cross-sell activos", "Sofia: de foco activos"], start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True, align="center" if i > 1 else None)
    f += 1
    rangos = ["1-10", "11-50", "51-250", "251-500", "501-1000", "1001+", "5000+"]
    foco = {n for n, _ in FABRICANTES_FOCO}
    for rango in rangos + ["Sin dato"]:
        e = [x for x in activos(equipo) if (x["empleados"] or "Sin dato") == rango]
        s_ = [x for x in activos(sofia) if (x["empleados"] or "Sin dato") == rango]
        fill = "F5F5F5" if rango == "Sin dato" else None
        celda(ws, f, 1, rango, fill=fill)
        for j, v in enumerate([len(e), sum(1 for x in e if x["tipo"] == "Cross-Sell"),
                               sum(1 for x in e if x["fabricante"] in foco),
                               len(s_), sum(1 for x in s_ if x["tipo"] == "Cross-Sell"),
                               sum(1 for x in s_ if x["fabricante"] in foco)], start=2):
            celda(ws, f, j, v, fill=fill, align="center")
        f += 1
    celda(ws, f, 1, "\"De foco\": Halo, Humand, Atera, NinjaOne y ServiceNow. Los rangos salen del campo "
          "Employee Count de la organizacion.", size=9, color="5B5F6B", borde=False)

    # --- Grafico: activos por tipo y vendedor ---
    from openpyxl.chart import BarChart, Reference
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = "bar", "stacked", 100
    ch.title = "Farmings activos por vendedor y tipo"
    ch.add_data(Reference(ws, min_col=4, max_col=7, min_row=5, max_row=5 + n_vend), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=6, max_row=5 + n_vend))
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.scaling.orientation = "maxMin"
    ch.y_axis.crosses = "max"
    ch.legend.position = "r"
    ch.height, ch.width = 7.5, 15
    ws.add_chart(ch, "I4")

    # --- C. Detalle de lo activo del equipo ---
    f += 2
    banda(ws, f, "C. Farmings activos del equipo (cross-sell primero, los mas antiguos arriba)", 11, C_SMB)
    f += 1
    cols = ["Farming", "Organizacion", "Vendedor", "Tipo", "Etapa", "Fabricante", "Producto", "Empleados",
            "Pais", "Creado", "Dias abiertos"]
    for i, t in enumerate(cols, start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)
    f += 1
    ini = f
    orden = {t: i for i, t in enumerate(TIPOS_FARMING)}
    hoy = date.today()
    # Se ordena por antiguedad y no por "Modified Time": en vTiger ese campo se
    # movio para los 44 farmings activos el mismo dia (27-10-2026), asi que no
    # mide actividad real del vendedor.
    for x in sorted(activos(equipo), key=lambda x: (orden.get(x["tipo"], 9), x["creado"] or date.max)):
        dias = (hoy - x["creado"]).days if x["creado"] else None
        fill = None
        for i, v in enumerate([x["farming"], x["org"], x["vendedor"], x["tipo"], x["etapa"],
                               x["fabricante"], x["producto"], x["empleados"], x["pais"],
                               f"{x['creado']:%d/%m/%Y}" if x["creado"] else "", dias], start=1):
            celda(ws, f, i, v, fill=fill, wrap=(i in (1, 2)), align="center" if i == 11 else None)
        f += 1
    if f == ini:
        celda(ws, f, 1, "El equipo no tiene farmings activos.")

    for i, ancho in enumerate([46, 30, 20, 18, 20, 16, 20, 12, 14, 12, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.sheet_view.showGridLines = False
    return ws


def hoja_okr_anual(wb, ctx):
    """Detalle deal por deal del acumulado del ano (YTD): de donde sale cada
    monto, incluyendo lo aportado por vendedores que ya no estan en el equipo
    (Legacy). Se llamaba "Detalle OKR Anual"; desde el 27-10-2026 ya no hay OKR
    y se renombro. Los deals de suscripcion de ServiceNow (los que cuentan para
    la meta de la seccion D) van resaltados."""
    ws = wb.create_sheet("Detalle Anual")
    ncols = 10
    celda(ws, 1, 1, f"Detalle del acumulado del año (YTD), al "
                    f"{ctx['hasta']:%d/%m/%Y}", bold=True, size=14, color=C_DARK, borde=False)
    celda(ws, 2, 1, "Todo lo ganado en el año, deal por deal. Los resaltados en azul son las "
                    "suscripciones de ServiceNow que cuentan para la meta de la seccion D del "
                    "Resumen Ejecutivo. Incluye al equipo actual y a vendedores que ya no estan "
                    "en GB Advisors (columna Fuente = Legacy), sumados aparte.",
          size=9, color="5B5F6B", wrap=True, borde=False)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    ws.row_dimensions[2].height = 30

    ytd = ctx["ytd_ganado"]
    nb = suma(por_categoria(ytd, "NEW LOGO"))
    exp = suma(por_categoria(ytd, "EXPANSION"))
    serv = suma(por_categoria(ytd, "SERVICES"))
    pf = suma(por_categoria(ytd, "PAYMENT FREQUENCY"))
    legacy = [d for d in ytd if d["fuente"] == "Legacy"]

    f = 4
    banda(ws, f, "Resumen (el de ServiceNow coincide con la seccion D del Resumen Ejecutivo)", ncols, C_SMB)
    f += 1
    sn = suma([d for d in ytd if es_suscripcion_servicenow(d)])
    for etiqueta, valor, negrita in [
        ("ServiceNow · Suscripciones (logos nuevos + expansion, sin renovaciones)", sn, True),
        ("Logos nuevos + expansion, todos los fabricantes", nb + exp, True),
        ("Servicios (incl. renovaciones de servicio)", serv, True),
        ("Payment Frequency", pf, True),
        ("De lo anterior, aportado por vendedores Legacy", suma(legacy), False),
    ]:
        celda(ws, f, 1, etiqueta, bold=negrita)
        c = celda(ws, f, 2, valor, fmt=MONEY, bold=negrita)
        ws.merge_cells(start_row=f, start_column=2, end_row=f, end_column=3)
        f += 1
    celda(ws, f, 1, f"{len(legacy)} deal(s) legacy de "
                    f"{', '.join(sorted({d['vendedor'] for d in legacy})) or 'ninguno'}.",
          size=9, color="5B5F6B", wrap=True, borde=False)
    ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=ncols)
    f += 2

    banda(ws, f, "Detalle deal por deal", ncols, C_SMB)
    f += 1
    cols = ["Deal ID", "Oportunidad", "Organizacion", "Vendedor", "Fuente", "Categoria",
            "Tipo", "ECD", "Monto CRM", "ARR"]
    for i, t in enumerate(cols, start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)
    f += 1
    for d in sorted(ytd, key=lambda x: (x["fuente"] == "Legacy", -x["arr"])):
        fill = ("FCE9E9" if d["fuente"] == "Legacy"
                else "DCE9F7" if es_suscripcion_servicenow(d) else None)
        for i, v in enumerate([d["deal_id"], d["nombre"], d["org"], d["vendedor"],
                               d["fuente"], d["categoria"], d["tipo"],
                               f"{d['ecd']:%d/%m/%Y}" if d["ecd"] else "",
                               d["mrr"], d["arr"]], start=1):
            c = celda(ws, f, i, v, fill=fill, wrap=(i == 2))
            if i == 9:
                c.number_format = MONEY2
            if i == 10:
                c.number_format = MONEY
        f += 1
    for i, ancho in enumerate([20, 44, 28, 20, 14, 15, 24, 12, 14, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = ws.cell(row=9, column=1)
    ws.sheet_view.showGridLines = False
    return ws


def hoja_calidad(wb, ctx):
    ws = wb.create_sheet("Calidad de Datos")
    celda(ws, 1, 1, f"Calidad de datos - deals del periodo ({ctx['periodo_label']})", bold=True,
          size=14, color=C_DARK, borde=False)
    celda(ws, 2, 1, "Solo se revisan los deals con ECD dentro del periodo: un hueco en un deal "
                    "de otro trimestre no afecta estas metricas. Corregirlos en vTiger.",
          size=9, color="5B5F6B", borde=False)
    for i, t in enumerate(["Hallazgo", "Deals", "Detalle"], start=1):
        celda(ws, 4, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID)

    deals = ctx["periodo"]
    checks = [
        ("Sin Deal ID", [d for d in deals if not d["deal_id"]]),
        ("Monto en cero o vacio", [d for d in deals if d["mrr"] == 0]),
        ("Monto negativo", [d for d in deals if d["mrr"] < 0]),
        ("Sin Lead Source", [d for d in deals if not d["lead_source"]]),
        ("Sin tipo (New/Renew/ProServ)", [d for d in deals if not d["tipo"]]),
        ("Categoria sin clasificar", [d for d in deals if d["categoria"] == "SIN CLASIFICAR"]),
    ]
    f = 5
    for etiqueta, grupo in checks:
        celda(ws, f, 1, etiqueta)
        c = celda(ws, f, 2, len(grupo), align="center", bold=True)
        if grupo:
            c.fill = PatternFill("solid", fgColor=C_YELLOW)
        ejemplos = ", ".join((d["deal_id"] or d["nombre"][:24]) for d in grupo[:6])
        celda(ws, f, 3, ejemplos + (" ..." if len(grupo) > 6 else ""), size=9,
              color="5B5F6B", wrap=True)
        f += 1

    # Los deals sin ECD no caen en ningun trimestre, asi que van aparte: no se
    # pueden acotar al periodo y quedarian invisibles.
    f += 1
    celda(ws, f, 1, "Aparte: deals sin fecha, no entran en ningun trimestre", bold=True,
          borde=False)
    f += 1
    sin_ecd = [d for d in ctx["equipo"] if not d["ecd"]]
    for etiqueta, grupo in (
        ("Sin Expected Close Date", sin_ecd),
        ("Ganado sin Expected Close Date", [d for d in sin_ecd if d["estado"] == "GANADO"]),
    ):
        celda(ws, f, 1, etiqueta)
        c = celda(ws, f, 2, len(grupo), align="center", bold=True)
        if grupo:
            c.fill = PatternFill("solid", fgColor=C_YELLOW)
        ejemplos = ", ".join((d["deal_id"] or d["nombre"][:24]) for d in grupo[:6])
        celda(ws, f, 3, ejemplos + (" ..." if len(grupo) > 6 else ""), size=9,
              color="5B5F6B", wrap=True)
        f += 1

    otros = sorted({d["vendedor"] for d in ctx["todos"]
                    if d["vendedor"] and d["vendedor"] not in EQUIPO})
    f += 1
    celda(ws, f, 1, "Nota de alcance", bold=True, borde=False)
    f += 1
    celda(ws, f, 1, f"El reporte cubre a los {len(EQUIPO)} vendedores del equipo: "
                    f"{', '.join(EQUIPO)}. Otros duenos que aparecen en el reporte y quedan "
                    f"fuera: {', '.join(otros) if otros else 'ninguno'}.",
          size=9, color="5B5F6B", wrap=True, borde=False)
    ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=3)
    ws.row_dimensions[f].height = 30

    for i, ancho in enumerate([36, 12, 90], start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.sheet_view.showGridLines = False
    return ws


# ---------- insights ----------

def construir_insights(ctx):
    out = []
    g = ctx["mes_ganado"]
    arr_mes, arr_prev = suma(g), suma(ctx["prev_ganado"])
    m, mp = ctx["periodo_label"], ctx["periodo_prev_label"]

    if arr_prev:
        var = (arr_mes - arr_prev) / arr_prev
        rumbo = "arriba" if var >= 0 else "abajo"
        out.append(f"{m}: ${arr_mes:,.0f} de ARR ganado, {abs(var):.0%} {rumbo} del "
                   f"trimestre anterior completo ({mp}, ${arr_prev:,.0f}).")
    else:
        out.append(f"{m}: ${arr_mes:,.0f} de ARR ganado.")

    exp = suma(por_categoria(g, "EXPANSION"))
    nb = suma(por_categoria(g, "NEW LOGO"))
    ventas_mes = exp + nb
    if ventas_mes:
        out.append(f"De los ${ventas_mes:,.0f} de logos nuevos y expansion del periodo, "
                   f"{exp / ventas_mes:.0%} vino de expansion sobre la base instalada y "
                   f"{nb / ventas_mes:.0%} de logos nuevos.")
    if nb == 0:
        out.append("ALERTA: no se ha cerrado ningun logo nuevo en el periodo; todo lo ganado "
                   "salio de la base instalada.")

    # ServiceNow: la unica meta definida (27-10-2026).
    es_mes = ctx["es_mes"]
    meta_per = TARGET_SERVICENOW_Q / 3 if es_mes else TARGET_SERVICENOW_Q
    sn_per = suma([d for d in g if es_suscripcion_servicenow(d)])
    esperado_per = meta_per * ctx["avance_meta"]
    out.append(f"ServiceNow, suscripciones del {'mes' if es_mes else 'trimestre'}: ${sn_per:,.0f} de "
               f"${meta_per:,.0f} ({sn_per / meta_per:.0%}) con {ctx['avance_meta']:.0%} del periodo "
               f"transcurrido; el ritmo esperado a hoy seria ${esperado_per:,.0f}.")

    ytd = ctx["ytd_ganado"]
    sn_ytd = suma([d for d in ytd if es_suscripcion_servicenow(d)])
    out.append(f"ServiceNow FY{str(ctx['anio_ref'])[-2:]}: ${sn_ytd:,.0f} de ${TARGET_SERVICENOW_ANIO:,.0f} "
               f"({sn_ytd / TARGET_SERVICENOW_ANIO:.1%}). Faltan ${max(TARGET_SERVICENOW_ANIO - sn_ytd, 0):,.0f}; "
               f"hay ${suma([d for d in ctx['abiertos'] if es_suscripcion_servicenow(d)]):,.0f} de ARR en "
               f"pipeline abierto de suscripciones de ServiceNow.")

    perdidos = ctx["mes_perdido"]
    wr = tasa(len(g), len(g) + len(perdidos))
    if wr is not None:
        out.append(f"Win rate del periodo: {wr:.0%} por cantidad ({len(g)} ganados, "
                   f"{len(perdidos)} perdidos), con ${suma(perdidos):,.0f} de ARR perdido.")

    top = max(g, key=lambda d: d["arr"], default=None)
    if top and arr_mes:
        conc = top["arr"] / arr_mes
        if conc >= 0.3:
            out.append(f"Riesgo de concentracion: {conc:.0%} del ARR del periodo viene de un solo "
                       f"deal ({top['nombre'][:52]}, {top['vendedor']}).")

    serv_ytd = suma(por_categoria(ytd, "SERVICES"))
    out.append(f"Servicios ganados en el ano: ${serv_ytd:,.0f} de ARR, incluyendo renovaciones de "
               f"servicio (sin meta definida).")

    pf_ytd = suma(por_categoria(ytd, "PAYMENT FREQUENCY"))
    pf_abierto = suma(por_categoria(ctx["abiertos"], "PAYMENT FREQUENCY"))
    out.append(f"Payment Frequency en el ano: ${pf_ytd:,.0f} ganados y ${pf_abierto:,.0f} abiertos "
               f"por cerrar (sin meta definida).")

    ch_si = [c for c in ctx["churn"] if c["corresponde"] == "Si"]
    if ch_si:
        out.append(f"Churn atribuible al equipo: ${sum(c['perdido'] for c in ch_si):,.2f} en "
                   f"{len(ch_si)} asset(s).")
    else:
        out.append(f"Sin churn atribuible al equipo en el periodo "
                   f"({len(ctx['churn'])} asset(s) revisados, ninguno aplica).")
    return out


# ---------- main ----------

def main():
    parser = argparse.ArgumentParser(
        description="Reporte ejecutivo de Mid/SMB. Por defecto el trimestre en curso al dia de hoy.")
    parser.add_argument("--trimestre", help="Trimestre completo, ej. 2026Q2.")
    parser.add_argument("--mes", help="Un mes especifico, ej. 2026-07.")
    args = parser.parse_args()

    hoy = date.today()
    if args.mes:
        anio, mes = int(args.mes[:4]), int(args.mes[5:7])
        desde, hasta = rango_mes(anio, mes)
        etiqueta = f"{MESES[mes - 1]} {anio}"
        p_anio, p_mes = mes_anterior(anio, mes)
        p_desde, p_hasta = rango_mes(p_anio, p_mes)
        etiqueta_prev = f"{MESES[p_mes - 1]} {p_anio}"
        anio_ref, sufijo = anio, f"{anio}-{mes:02d} {MESES[mes - 1]}"
    else:
        if args.trimestre:
            anio, q = int(args.trimestre[:4]), int(args.trimestre[-1])
            desde, hasta = rango_trimestre(anio, q)
            etiqueta = f"Q{q} {anio}"
        else:
            anio, q = hoy.year, trimestre_de(hoy)
            desde, _ = rango_trimestre(anio, q)
            hasta = hoy  # hasta el momento de ejecutar
            etiqueta = f"Q{q} {anio} al {hoy:%d/%m/%Y}"
        p_anio, p_q = trimestre_anterior(anio, q)
        p_desde, p_hasta = rango_trimestre(p_anio, p_q)
        etiqueta_prev = f"Q{p_q} {p_anio}"
        anio_ref, sufijo = anio, f"{anio}Q{q}"

    reporte = find_reporte("*DIQ*.xlsx", "DIQ")
    log(f"Reporte: {reporte.name}")
    edad = (datetime.now() - datetime.fromtimestamp(reporte.stat().st_mtime)).days
    if edad >= 1:
        log(f"AVISO: el reporte tiene {edad} dia(s); descargalo de nuevo para datos frescos.")
    header, filas = leer(reporte)
    todos = normalizar(filas)
    equipo = [d for d in todos if d["vendedor"] in EQUIPO]
    log(f"Deals del equipo en el reporte: {len(equipo)} de {len(todos)}")

    periodo_deals = en_rango(equipo, desde, hasta)
    prev_deals = en_rango(equipo, p_desde, p_hasta)
    ytd_deals = en_rango(equipo, date(anio_ref, 1, 1), hasta)

    # Vendedores que ya no estan en el equipo (ej. Fares Rivas, Antonio
    # Morales) no aparecen en el DIQ actual, asi que su ARR ganado queda
    # afuera del OKR anual aunque cuente para la meta de la empresa. Un
    # reporte aparte (bajado a mano, nombre con "Legacy") los suma solo al
    # YTD del OKR anual (seccion D) -- no entran a "equipo"/EQUIPO ni afectan
    # el avance del trimestre, la hoja Por Vendedor ni la nota de alcance.
    r_legacy = find_reporte("*Legacy*.xlsx", "Legacy SR", obligatorio=False)
    legacy_ytd_ganado = []
    if r_legacy:
        _, legacy_filas = leer(r_legacy)
        legacy_ytd = en_rango(normalizar(legacy_filas), date(anio_ref, 1, 1), hasta)
        legacy_ytd_ganado = [d for d in legacy_ytd if d["estado"] == "GANADO"]
        log(f"Legacy SR: {len(legacy_ytd_ganado)} deals ganados YTD "
            f"(${suma(legacy_ytd_ganado):,.0f} ARR) sumados al acumulado anual.")
    else:
        log("AVISO: falta el reporte Legacy SR; el acumulado anual no incluye a "
            "vendedores que ya no estan en el equipo.")

    # Churn: se reutiliza el analisis del plan de compensacion.
    spec = importlib.util.spec_from_file_location("pc", SCRIPT_DIR / "Plan Compensacion.py")
    pc = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(pc)
    r_churn = find_reporte("*ChurnAssets*.xlsx", "ChurnAssets", obligatorio=False)
    r_oa = find_reporte("*- OA_*.xlsx", "OA", obligatorio=False)
    churn = []
    if r_churn and r_oa:
        _, cf = leer(r_churn)
        _, of = leer(r_oa)
        # Acotar al periodo igual que el plan de compensacion: el export de
        # ChurnAssets trae todo lo cargado en vTiger, no solo el trimestre, y
        # sin este filtro la seccion de churn de un Q4 mostraba los 9 churns de
        # Q3 (1-oct-2026).
        cf, fuera = pc.churn_del_trimestre(cf, desde, hasta)
        if fuera:
            log(f"Churn: {fuera} asset(s) fuera del periodo, no se cuentan.")
        churn = pc.analizar_churn(cf, of, header, filas)
    else:
        log("AVISO: falta ChurnAssets u OA; la seccion de churn queda vacia.")

    # Ritmo dentro del periodo: que fraccion del trimestre ya transcurrio.
    dias_periodo = (rango_trimestre(anio, trimestre_de(desde))[1] - desde).days + 1
    dias_corridos = (hasta - desde).days + 1
    # Avance para la meta de ServiceNow: se mide hasta HOY aunque el periodo
    # pedido termine despues (un --mes o --trimestre del periodo en curso).
    fecha_ref = min(hasta, hoy)
    dias_meta = (hasta - desde).days + 1 if args.mes else dias_periodo
    dias_corridos_meta = max(min((fecha_ref - desde).days + 1, dias_meta), 0)
    avance_meta = dias_corridos_meta / dias_meta if dias_meta else 1.0
    ini_qr, fin_qr = rango_trimestre(anio_ref, trimestre_de(fecha_ref))
    avance_trimestre_actual = min(((fecha_ref - ini_qr).days + 1) / ((fin_qr - ini_qr).days + 1), 1.0)
    # Farming: Clauco Org (todo GB, con fecha de creacion). Opcional: sin el
    # reporte la hoja avisa que falta en vez de romper el resto.
    r_far = find_reporte("Clauco Org*.xlsx", "Clauco Org", obligatorio=False)
    farmings = cargar_farmings(r_far) if r_far else None
    farmings_fuente = f"{r_far.stem.split('_')[0]} del {datetime.fromtimestamp(r_far.stat().st_mtime):%d/%m}" if r_far else ""
    if r_far:
        log(f"Farming: {len(farmings)} farmings de {r_far.name}")

    ctx = {
        "reporte": reporte,
        "periodo_label": etiqueta,
        "periodo_prev_label": etiqueta_prev,
        "desde": desde, "hasta": hasta,
        "meses_transcurridos": hasta.month if anio_ref == hoy.year else 12,
        "avance_periodo": min(dias_corridos / dias_periodo, 1.0) if dias_periodo else 1.0,
        "dias_corridos": dias_corridos, "dias_periodo": dias_periodo,
        "es_mes": bool(args.mes), "anio_ref": anio_ref, "fecha_ref": fecha_ref,
        "avance_meta": avance_meta, "dias_meta": dias_meta, "dias_corridos_meta": dias_corridos_meta,
        "avance_trimestre_actual": avance_trimestre_actual,
        "todos": todos, "equipo": equipo, "periodo": periodo_deals,
        "mes_ganado": [d for d in periodo_deals if d["estado"] == "GANADO"],
        "mes_perdido": [d for d in periodo_deals if d["estado"] == "PERDIDO"],
        "prev_ganado": [d for d in prev_deals if d["estado"] == "GANADO"],
        "ytd_ganado": _con_fuente([d for d in ytd_deals if d["estado"] == "GANADO"], "Equipo actual")
                      + _con_fuente(legacy_ytd_ganado, "Legacy"),
        "abiertos": [d for d in equipo if d["estado"] == "ABIERTO"],
        "churn": churn,
        "farmings": farmings, "farmings_fuente": farmings_fuente,
    }
    ctx["insights"] = construir_insights(ctx)

    log(f"Periodo: {etiqueta} ({desde} a {hasta}) | ganados {len(ctx['mes_ganado'])} "
        f"(${suma(ctx['mes_ganado']):,.0f} ARR) | perdidos {len(ctx['mes_perdido'])} "
        f"| pipeline abierto ${suma(ctx['abiertos']):,.0f}")

    wb = Workbook()
    wb.remove(wb.active)
    hoja_resumen(wb, ctx)
    hoja_vendedores(wb, ctx)
    hoja_pipeline(wb, ctx)
    hoja_deals(wb, ctx)
    hoja_farming(wb, ctx)
    hoja_okr_anual(wb, ctx)
    hoja_calidad(wb, ctx)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    # Nombre estable por periodo: cada clic sobrescribe el mismo archivo.
    salida = OUTPUT_DIR / f"Iran team metrics {sufijo}.xlsx"
    try:
        wb.save(salida)
    except PermissionError:
        raise RuntimeError(
            f"No se pudo escribir '{salida.name}' porque esta abierto en Excel. "
            f"Cierralo y vuelve a generar."
        ) from None
    log(f"Listo: {salida}")
    return salida




if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        # Con el tipo y la traza: un StopIteration (o cualquier excepcion sin
        # mensaje) dejaba un "ERROR: " vacio que no decia nada (1-oct-2026).
        log(f"ERROR [{type(exc).__name__}]: {exc or 'sin mensaje'}")
        traceback.print_exc()
        sys.exit(1)
