"""
GBS Deal Checker: revisa que cada deal ABIERTO del equipo (Deals Sales Stage
distinto de Closed Won/Closed Lost) tenga un "Deals GBS Link" en el ultimo
DIQ descargado de vTiger, y que ese link responda -- no un 404/403 directo,
que en SharePoint suele significar que el documento detras ya no existe, sin
importar quien lo abra (hallazgo del 28-ago-2026: los links personales de
gente que ya salio del tenant -Antonio Morales, Sandra Agudo, Oscar Andrade,
Nohelia Carrillo- quedan muertos para siempre).

De solo lectura hacia vTiger: no escribe nada, no manda correos. Guarda un
Excel resumen en "GBS Deal Checker/informes/gbs-checker-YYYY-MM-DD.xlsx".

Ejecutar con: python gbs_deal_checker.py
"""
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlsplit, urlunsplit
from urllib.request import Request, urlopen

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import generar_reportes_next_step as gen_ns  # reutiliza descubrimiento/lectura/reglas ya probadas

from paths import GB_CLAUDE, GBS_CHECKER_INFORMES_DIR as INFORMES_DIR

TIMEOUT_SEG = 12
MAX_WORKERS = 16

ORDEN_ESTADO = {"SIN_LINK": 0, "ROTO": 1, "NO_VERIFICABLE": 2, "OTRO": 3, "OK": 4}


def normalizar_url(url):
    """Codifica los caracteres que una URL no puede llevar crudos, sobre todo
    los espacios.

    Hasta el 24-sep-2026 cualquier link con un espacio se descartaba como
    NO_VERIFICABLE sin llegar a probarlo, y eso marcaba como sospechosos
    links perfectamente validos: SharePoint guarda rutas como
    ".../sites/sales/Shared Documents/<Cliente con espacios>/<archivo>.docx",
    que es exactamente la forma que tienen los GBS de este equipo. Caso real:
    el deal de Anglican Diocese salia NO_VERIFICABLE y su documento abria sin
    problema (responde 200 una vez codificado).

    Un espacio no es un link corrupto, es un link sin codificar."""
    p = urlsplit(url)
    return urlunsplit((
        p.scheme, p.netloc,
        quote(p.path, safe="/%:@&=+$,~"),
        quote(p.query, safe="=&%:/?+,~"),
        p.fragment,
    ))


def check_link(url):
    """OK si responde con normalidad (200, o redirect que termina ahi, tipico
    de un link valido sin sesion iniciada). ROTO si el servidor devuelve
    404/410 directo -- eso SI es una senal fuerte de que el documento ya no
    existe, con o sin sesion. NO_VERIFICABLE si el link no es una URL (no
    empieza por http) o hubo un error de red."""
    if not url or not str(url).strip():
        return "SIN_LINK"
    url = str(url).strip()
    if not url.lower().startswith("http"):
        return "NO_VERIFICABLE"
    try:
        url = normalizar_url(url)
    except Exception:
        return "NO_VERIFICABLE"
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(req, timeout=TIMEOUT_SEG) as resp:
            return "OK" if resp.status < 400 else f"OTRO ({resp.status})"
    except HTTPError as exc:
        if exc.code in (404, 410):
            return "ROTO"
        return f"OTRO ({exc.code})"
    except URLError:
        return "NO_VERIFICABLE"
    except Exception:
        return "NO_VERIFICABLE"


def load_open_deals():
    """Deals abiertos del equipo (TEAM_OWNERS), unicos por Deal ID, tal como
    figuran en el DIQ mas reciente de Descargas."""
    files = gen_ns.find_latest_files()
    diq_path = files.get("diq")
    if diq_path is None:
        return None, None

    rows = gen_ns.read_xlsx_rows(diq_path)
    seen_ids = set()
    deals = []
    for r in rows:
        deal_id = r.get("Deals Deal ID")
        if not deal_id or deal_id in seen_ids:
            continue
        owner = gen_ns.canonical_owner(r.get("Deals Assigned To"))
        if owner not in gen_ns.TEAM_OWNERS:
            continue
        if gen_ns.is_closed_deal(r):
            continue
        seen_ids.add(deal_id)
        deals.append({
            "owner": owner,
            "deal": str(r.get("Deals Deal Name") or "").strip(),
            "org": str(r.get("Deals Organization Name") or "").strip(),
            "stage": str(r.get("Deals Sales Stage") or "").strip(),
            "link": r.get("Deals GBS Link"),
        })
    return deals, diq_path


def main():
    deals, diq_path = load_open_deals()
    if deals is None:
        print("No se encontró el reporte DIQ más reciente en Descargas. Abortando.")
        sys.exit(1)

    print(f"Revisando GBS Link de {len(deals)} deal(es) abierto(s) del equipo "
          f"(fuente: {diq_path.name})…")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futuros = {ex.submit(check_link, d["link"]): d for d in deals}
        for fut in as_completed(futuros):
            futuros[fut]["estado"] = fut.result()

    deals.sort(key=lambda d: (ORDEN_ESTADO.get(d["estado"], 3), d["owner"], d["deal"]))

    conteos = {}
    for d in deals:
        conteos[d["estado"]] = conteos.get(d["estado"], 0) + 1

    INFORMES_DIR.mkdir(parents=True, exist_ok=True)
    hoy = datetime.now().strftime("%Y-%m-%d")
    salida = INFORMES_DIR / f"gbs-checker-{hoy}.xlsx"

    con_problema_lista = [d for d in deals if d["estado"] != "OK"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "GBS Deal Checker"
    ws.append(["Representante", "Deal", "Organización", "Etapa", "Estado GBS Link", "Link"])
    for d in con_problema_lista:
        ws.append([d["owner"], d["deal"], d["org"], d["stage"], d["estado"], d["link"] or ""])
    for col, ancho in zip("ABCDEF", (18, 40, 32, 16, 16, 70)):
        ws.column_dimensions[col].width = ancho
    # Si el Excel del dia esta abierto en Excel, Windows lo bloquea y save()
    # lanza PermissionError DESPUES de haber hecho todos los chequeos de red,
    # que son la parte cara. En vez de perder la corrida entera, se guarda al
    # lado con un sufijo y se avisa (24-sep-2026, le paso a el responsable comercial).
    try:
        wb.save(salida)
    except PermissionError:
        n = 2
        while (INFORMES_DIR / f"gbs-checker-{hoy}-{n}.xlsx").exists():
            n += 1
        alterna = INFORMES_DIR / f"gbs-checker-{hoy}-{n}.xlsx"
        wb.save(alterna)
        print(f"AVISO: '{salida.name}' esta abierto en Excel y no se pudo sobrescribir.")
        print(f"       El resultado se guardo en '{alterna.name}'. Cerra el archivo y, si "
              f"queres, renombralo.")
        salida = alterna

    print(f"Guardado: {salida} ({len(con_problema_lista)} fila(s), solo los que tienen "
          f"problema -- los OK no se listan)")
    print(f"Total deals abiertos revisados: {len(deals)}")
    print(f"  Sin GBS Link: {conteos.get('SIN_LINK', 0)}")
    print(f"  Link roto (404/410): {conteos.get('ROTO', 0)}")
    print(f"  No verificable (formato/red): {conteos.get('NO_VERIFICABLE', 0)}")
    print(f"  OK (no se incluyen en el Excel): {conteos.get('OK', 0)}")
    if con_problema_lista:
        print(f"Alerta: {len(con_problema_lista)} deal(es) abierto(s) sin un GBS Link válido.")
    else:
        print("Todos los deals abiertos tienen un GBS Link válido.")


if __name__ == "__main__":
    main()
