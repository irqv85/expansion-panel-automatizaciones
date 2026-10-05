"""
Reporte ejecutivo del departamento Mid/SMB, para compartir hacia arriba
(nivel C). Por defecto cubre el trimestre en curso hasta el momento de
ejecutarlo (QTD).

Lee los reportes que ya descargas (DIQ, ChurnAssets, OA) y produce un Excel
con el periodo, comparado contra el trimestre anterior completo, la meta
trimestral y los OKR anuales 2026.

Reglas financieras (OKR 2026 GB Advisors):
  - ARR: los servicios (ProfessionalServices / SSI add on) YA son valor anual,
    asi que ARR = monto. Todo lo demas es MRR mensual: ARR = monto x 12.
  - Categoria, segun el campo New/Renew/ProServ:
    "New" = NEW LOGO · "New-Add-On" y "New-Cross-Sell" = EXPANSION (caen sobre
    cuentas que ya son clientes) · "Payment Frequency" = su propia categoria,
    porque es el KR2 y sumarlo a Expansion contaria el mismo ARR dos veces ·
    ProfessionalServices, "SSI add on" y "Renew SSI" = SERVICES (el KR3 dice
    "incl. renewals") · "Renewal" de licencia = RENEWAL, es del equipo de
    Renewals.
  - Ganado = Closed Won + PO Invoiced. Perdido = Closed Lost. Abierto = resto.
  - Target de pipeline dinamico: (target de ventas - ARR ganado YTD) / 39%.

Uso:
  python "Metricas Mes.py"                    Trimestre en curso al dia de hoy
  python "Metricas Mes.py" --trimestre 2026Q2  Un trimestre completo
  python "Metricas Mes.py" --mes 2026-07       Un mes especifico
"""
import argparse
import calendar
import importlib.util
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

# ---- OKR 2026: metas del equipo, ver equipo.json ----
# Metas tal como estan en la hoja de OKR del equipo.
TARGET_SALES_ARR = 2_070_000   # KR1 Total FW Mid + Digital ARR (NB + Expansion)
TARGET_PF_ARR = 300_000        # KR2 Reseller Conversion: Payment Frequency
TARGET_SERVICES_ARR = 150_000  # KR3 Professional Services Revenue (incl. renewals)
TARGET_LOGOS = 20              # KR4 New partner logos (HALO + Humand + Intercom)
# Contexto del OKR (dato de la hoja, no se calcula de los reportes)
RESELLER_UNIVERSE = 66         # orgs Mfr Monthly 51-500 empleados
CONVERTIBLE_ARR = 955_753
CONVERTIBLE_TARGET = 382_301   # 40% del ARR convertible
CONV_RATE = 0.39

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

    # --- C. Avance del trimestre vs meta trimestral ---
    banda(ws, f, "C. AVANCE DEL TRIMESTRE VS META TRIMESTRAL", ncols, C_SMB)
    f += 1
    for i, t in enumerate(["Metrica", "Ganado en el periodo", "Meta del trimestre",
                           "% de la meta", "Ritmo esperado a hoy", "Estado"], start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)
    f += 1

    g_per = ctx["mes_ganado"]
    nb_per = suma(por_categoria(g_per, "NEW LOGO"))
    exp_per = suma(por_categoria(g_per, "EXPANSION"))
    serv_per = suma(por_categoria(g_per, "SERVICES"))
    avance = ctx["avance_periodo"]

    pf_per = suma(por_categoria(g_per, "PAYMENT FREQUENCY"))
    filas_q = [
        ("KR1 · FW Mid + Digital ARR (NB + Expansion)", nb_per + exp_per, TARGET_SALES_ARR / 4),
        ("KR2 · Reseller conversion: Payment Frequency", pf_per, TARGET_PF_ARR / 4),
        ("KR3 · Professional Services (incl. renewals)", serv_per, TARGET_SERVICES_ARR / 4),
    ]
    for etiqueta, real, meta_q in filas_q:
        pct = tasa(real, meta_q)
        esperado = meta_q * avance
        celda(ws, f, 1, etiqueta)
        celda(ws, f, 2, real, fmt=MONEY, bold=True)
        c = celda(ws, f, 3, meta_q, fmt=MONEY)
        c.comment = Comment("Meta anual del OKR dividida entre 4 trimestres.", "Metricas Mes")
        cp = celda(ws, f, 4, f"=B{f}/C{f}", fmt=PCT, bold=True)
        cp.fill = PatternFill("solid", fgColor=color_semaforo(pct))
        cp.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        c2 = celda(ws, f, 5, esperado, fmt=MONEY, color="5B5F6B")
        c2.comment = Comment(f"Meta del trimestre x {avance:.0%} del trimestre transcurrido "
                             f"({ctx['dias_corridos']} de {ctx['dias_periodo']} dias).",
                             "Metricas Mes")
        celda(ws, f, 6, "Sobre ritmo" if real >= esperado else "Bajo el ritmo esperado", size=9,
              color=("1F5C2E" if real >= esperado else C_RED), wrap=True)
        f += 1
    celda(ws, f, 1, f"Transcurrido {avance:.0%} del trimestre ({ctx['dias_corridos']} de "
                    f"{ctx['dias_periodo']} dias).", size=9, color="5B5F6B", wrap=True)
    ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=ncols)
    f += 2

    # --- D. Avance vs OKR 2026 ---
    banda(ws, f, "D. AVANCE ANUAL VS OKR 2026 (acumulado del año)", ncols, C_SMB)
    f += 1
    for i, t in enumerate(["Metrica", "YTD real", "Target 2026", "% avance",
                           "Ritmo esperado", "Estado"], start=1):
        celda(ws, f, i, t, bold=True, size=9, color="FFFFFF", fill=C_MID, wrap=True)
    f += 1

    ytd = ctx["ytd_ganado"]
    nb_ytd = suma(por_categoria(ytd, "NEW LOGO"))
    exp_ytd = suma(por_categoria(ytd, "EXPANSION"))
    serv_ytd = suma(por_categoria(ytd, "SERVICES"))
    ventas_ytd = nb_ytd + exp_ytd
    ritmo = ctx["meses_transcurridos"] / 12

    pf_ytd = suma(por_categoria(ytd, "PAYMENT FREQUENCY"))
    filas_okr = [
        ("KR1 · FW Mid + Digital ARR (NB + Expansion)", ventas_ytd, TARGET_SALES_ARR),
        ("KR2 · Reseller conversion: Payment Frequency", pf_ytd, TARGET_PF_ARR),
        ("KR3 · Professional Services (incl. renewals)", serv_ytd, TARGET_SERVICES_ARR),
    ]
    for etiqueta, real, target in filas_okr:
        pct = tasa(real, target)
        esperado = target * ritmo
        celda(ws, f, 1, etiqueta)
        celda(ws, f, 2, real, fmt=MONEY, bold=True)
        celda(ws, f, 3, target, fmt=MONEY)
        c = celda(ws, f, 4, f"=B{f}/C{f}", fmt=PCT, bold=True)
        c.fill = PatternFill("solid", fgColor=color_semaforo(pct))
        c.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        celda(ws, f, 5, esperado, fmt=MONEY, color="5B5F6B")
        estado = "Sobre ritmo" if real >= esperado else "Bajo el ritmo esperado"
        celda(ws, f, 6, estado, size=9,
              color=("1F5C2E" if real >= esperado else C_RED), wrap=True)
        f += 1

    # KR4: los reportes locales no traen el partner del deal, asi que no se
    # puede contar logos sin inventar el dato.
    celda(ws, f, 1, "KR4 · New partner logos (HALO + Humand + Intercom)", color="5B5F6B")
    c = celda(ws, f, 2, "—", color="5B5F6B", fill=C_GREY, align="center")
    c.comment = Comment("El DIQ no trae el partner ni el producto del deal, de modo que no se "
                        "puede contar logos con los reportes actuales. Haria falta agregar "
                        "Products Manufacturer / Product Name al reporte.", "Metricas Mes")
    celda(ws, f, 3, TARGET_LOGOS, color="5B5F6B", fill=C_GREY, align="center")
    celda(ws, f, 4, "sin medir", size=9, color="5B5F6B", fill=C_GREY, wrap=True)
    celda(ws, f, 5, None, fill=C_GREY)
    celda(ws, f, 6, "Falta el campo de partner en el reporte", size=9, color="5B5F6B", wrap=True)
    f += 1

    # Pipeline con target dinamico
    gap = max(TARGET_SALES_ARR - ventas_ytd, 0)
    pipe_target = round(gap / CONV_RATE) if gap > 0 else 0
    pipe_real = suma(ctx["abiertos"])
    pct_pipe = tasa(pipe_real, pipe_target) if pipe_target else None
    celda(ws, f, 1, "Pipeline ARR activo")
    celda(ws, f, 2, pipe_real, fmt=MONEY, bold=True)
    c = celda(ws, f, 3, pipe_target, fmt=MONEY)
    c.comment = Comment("Calculado dinamicamente: (target de ventas - ARR ganado YTD) / 39% "
                        "de conversion.", "Metricas Mes")
    if pipe_target:
        cp = celda(ws, f, 4, f"=B{f}/C{f}", fmt=PCT, bold=True)
        cp.fill = PatternFill("solid", fgColor=color_semaforo(pct_pipe))
        cp.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    else:
        celda(ws, f, 4, "Target alcanzado", size=9, fill=C_GREEN, color="FFFFFF", wrap=True)
    celda(ws, f, 5, None)
    celda(ws, f, 6, f"Cobertura necesaria: {gap / CONV_RATE / 1000:,.0f}K" if gap else "—",
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


def hoja_okr_anual(wb, ctx):
    """Detalle deal por deal de la seccion D (Avance Anual vs OKR): de donde
    sale cada monto que compone el YTD, incluyendo lo aportado por vendedores
    que ya no estan en el equipo (Legacy)."""
    ws = wb.create_sheet("Detalle OKR Anual")
    ncols = 10
    celda(ws, 1, 1, f"Detalle del acumulado del año (YTD) para el OKR 2026, al "
                    f"{ctx['hasta']:%d/%m/%Y}", bold=True, size=14, color=C_DARK, borde=False)
    celda(ws, 2, 1, "Es el mismo total de la seccion D del Resumen Ejecutivo, deal por deal. "
                    "Incluye al equipo actual y a vendedores que ya no estan en GB Advisors "
                    "(columna Fuente = Legacy), sumados aparte porque su ARR ya cuenta para "
                    "la meta anual de la empresa.",
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
    banda(ws, f, "Resumen (coincide con la seccion D del Resumen Ejecutivo)", ncols, C_SMB)
    f += 1
    for etiqueta, valor, negrita in [
        ("KR1 · FW Mid + Digital ARR (NB + Expansion)", nb + exp, True),
        ("KR2 · Reseller conversion: Payment Frequency", pf, True),
        ("KR3 · Professional Services (incl. renewals)", serv, True),
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
        fill = "FCE9E9" if d["fuente"] == "Legacy" else None
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
        out.append(f"De los ${ventas_mes:,.0f} del KR1 en el periodo, {exp / ventas_mes:.0%} "
                   f"vino de expansion sobre la base instalada y {nb / ventas_mes:.0%} de logos "
                   f"nuevos.")
    if nb == 0:
        out.append("ALERTA: no se ha cerrado ningun logo nuevo en el periodo; todo el KR1 "
                   "salio de la base instalada.")

    meta_q = TARGET_SALES_ARR / 4
    ventas_per = exp + nb
    esperado_q = meta_q * ctx["avance_periodo"]
    out.append(f"Contra la meta del trimestre (${meta_q:,.0f}), el equipo va en "
               f"${ventas_per:,.0f} ({ventas_per / meta_q:.0%}) con "
               f"{ctx['avance_periodo']:.0%} del trimestre transcurrido; el ritmo esperado a hoy "
               f"seria ${esperado_q:,.0f}.")

    ytd = ctx["ytd_ganado"]
    ventas_ytd = suma(por_categoria(ytd, "NEW LOGO")) + suma(por_categoria(ytd, "EXPANSION"))
    pct = ventas_ytd / TARGET_SALES_ARR
    ritmo = ctx["meses_transcurridos"] / 12
    esperado = TARGET_SALES_ARR * ritmo
    dif = ventas_ytd - esperado
    out.append(f"KR1 YTD: ${ventas_ytd:,.0f} de ${TARGET_SALES_ARR:,.0f} "
               f"({pct:.1%}). Al mes {ctx['meses_transcurridos']} el ritmo esperado seria "
               f"${esperado:,.0f}: hay una brecha de ${abs(dif):,.0f} "
               f"{'a favor' if dif >= 0 else 'en contra'}.")

    gap = max(TARGET_SALES_ARR - ventas_ytd, 0)
    if gap:
        necesario = gap / CONV_RATE
        pipe = suma(ctx["abiertos"])
        cobertura = pipe / necesario if necesario else None
        out.append(f"Para cerrar la brecha hacen falta ${necesario:,.0f} de pipeline al 39% "
                   f"de conversion. Hoy hay ${pipe:,.0f} abiertos, o sea "
                   f"{cobertura:.0%} de la cobertura necesaria.")

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
    out.append(f"KR3 Servicios YTD: ${serv_ytd:,.0f} de ${TARGET_SERVICES_ARR:,.0f} "
               f"({serv_ytd / TARGET_SERVICES_ARR:.0%}), incluyendo renovaciones de servicio.")

    pf_ytd = suma(por_categoria(ytd, "PAYMENT FREQUENCY"))
    pf_abierto = suma(por_categoria(ctx["abiertos"], "PAYMENT FREQUENCY"))
    if pf_ytd == 0:
        out.append(f"ALERTA KR2: Payment Frequency sigue en $0 ganado de "
                   f"${TARGET_PF_ARR:,.0f}. Hay ${pf_abierto:,.0f} en deals abiertos de ese "
                   f"tipo, sobre un ARR convertible de ${CONVERTIBLE_ARR:,.0f} en "
                   f"{RESELLER_UNIVERSE} orgs (la meta del 40% son ${CONVERTIBLE_TARGET:,.0f}).")
    else:
        out.append(f"KR2 Payment Frequency: ${pf_ytd:,.0f} de ${TARGET_PF_ARR:,.0f} "
                   f"({pf_ytd / TARGET_PF_ARR:.0%}), con ${pf_abierto:,.0f} abiertos por cerrar.")

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
            f"(${suma(legacy_ytd_ganado):,.0f} ARR) sumados al OKR anual.")
    else:
        log("AVISO: falta el reporte Legacy SR; el OKR anual no incluye a "
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
    ctx = {
        "reporte": reporte,
        "periodo_label": etiqueta,
        "periodo_prev_label": etiqueta_prev,
        "desde": desde, "hasta": hasta,
        "meses_transcurridos": hasta.month if anio_ref == hoy.year else 12,
        "avance_periodo": min(dias_corridos / dias_periodo, 1.0) if dias_periodo else 1.0,
        "dias_corridos": dias_corridos, "dias_periodo": dias_periodo,
        "todos": todos, "equipo": equipo, "periodo": periodo_deals,
        "mes_ganado": [d for d in periodo_deals if d["estado"] == "GANADO"],
        "mes_perdido": [d for d in periodo_deals if d["estado"] == "PERDIDO"],
        "prev_ganado": [d for d in prev_deals if d["estado"] == "GANADO"],
        "ytd_ganado": _con_fuente([d for d in ytd_deals if d["estado"] == "GANADO"], "Equipo actual")
                      + _con_fuente(legacy_ytd_ganado, "Legacy"),
        "abiertos": [d for d in equipo if d["estado"] == "ABIERTO"],
        "churn": churn,
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
    hoja_okr_anual(wb, ctx)
    hoja_calidad(wb, ctx)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    # Nombre estable por periodo: cada clic sobrescribe el mismo archivo.
    salida = OUTPUT_DIR / f"Metricas Mid-SMB {sufijo}.xlsx"
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
