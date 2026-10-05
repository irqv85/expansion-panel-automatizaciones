"""
Calcula que organizaciones del equipo necesitan un GBS nuevo: las que estan
asignadas a un vendedor del equipo (TEAM_OWNERS) y no tienen ningun
"Organizations GBS Link" en el ultimo ORGIQ descargado de vTiger, y que
todavia no tienen ni un borrador ni un GBS ya desplegado para esa misma
organizacion. Escribe el plan en plan_creacion_gbs_orgs.json. No toca
vTiger ni SharePoint -- eso lo hace Claude Code despues, siguiendo la
seccion GENERAR del skill "gbs-organization-maker" con este plan ya
calculado (mismo patron "Calcular -> Ejecutar" que el resto del panel).

A diferencia de los deals, una organizacion no tiene "Sales Stage" -- no
hay concepto de abierta/cerrada a este nivel, asi que el unico filtro es
que este asignada al equipo y sin GBS Link.

LIMITE POR CORRIDA (--limite N, pedido de el responsable comercial el 28-ago-2026): el lote
completo son ~150 organizaciones y cada una consume creditos de Claude en
la corrida de GENERAR, asi que el panel permite procesarlas por tandas.
Cuando se pasa --limite, el plan se recorta a las N PRIMERAS de una cola
ORDENADA POR CANTIDAD DE DEALS ASOCIADOS (descendente): una organizacion
sin ningun deal no tiene de donde sacar pain points y saldria "Sin pain
points registrados", asi que no tiene sentido que gaste un lugar de la
tanda -- queda al final de la cola. (Dato del 28-ago-2026: de 154
organizaciones sin GBS Link, 64 no tienen ningun deal asociado.)

Ejecutar con: python planificar_creacion_gbs_orgs.py [--limite N]
"""
import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import generar_reportes_next_step as gen_ns  # reutiliza find_latest_files/read_xlsx_rows/TEAM_OWNERS

from paths import GBS_ORG_MAKER_DIR as MAKER_DIR

SCRIPT_DIR = Path(__file__).resolve().parent
BORRADORES_DIR = MAKER_DIR / "borradores"
DESPLEGADOS_DIR = MAKER_DIR / "desplegados"
PLAN_FILE = SCRIPT_DIR / "plan_creacion_gbs_orgs.json"


def nombre_archivo(org):
    base = re.sub(r'[\\/:*?"<>|]', "-", str(org).strip())
    return base[:180] + ".docx"


def contar_deals_por_org():
    """Cuantos deals distintos tiene cada organizacion en el DIQ mas
    reciente. Sirve para priorizar la cola: una organizacion sin ningun
    deal asociado no tiene de donde sacar pain points (el skill la marcaria
    "Sin pain points registrados" sin generar nada), asi que no deberia
    gastar un lugar de una tanda limitada."""
    files = gen_ns.find_latest_files()
    diq_path = files.get("diq")
    if diq_path is None:
        return {}
    por_org = defaultdict(set)
    for r in gen_ns.read_xlsx_rows(diq_path):
        org = str(r.get("Deals Organization Name") or "").strip()
        deal_id = r.get("Deals Deal ID")
        if org and deal_id:
            por_org[org].add(deal_id)
    return {org: len(ids) for org, ids in por_org.items()}


def cargar_orgs_sin_link():
    files = gen_ns.find_latest_files()
    orgiq_path = files.get("orgiq")
    if orgiq_path is None:
        return None, None

    rows = gen_ns.read_xlsx_rows(orgiq_path)
    vistos = set()
    filas = []
    for r in rows:
        org = str(r.get("Organizations Organization Name") or "").strip()
        if not org or org in vistos:
            continue
        owner = gen_ns.canonical_owner(r.get("Organizations Assigned To"))
        if owner not in gen_ns.TEAM_OWNERS:
            continue
        link = r.get("Organizations GBS Link")
        if link and str(link).strip():
            continue
        vistos.add(org)
        filas.append({"owner": owner, "org": org})
    return filas, orgiq_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limite", type=int, default=None,
                    help="Máximo de organizaciones a incluir en esta corrida (las que "
                         "más deals tienen van primero). Sin este flag, entran todas.")
    args = ap.parse_args()

    filas, orgiq_path = cargar_orgs_sin_link()
    if filas is None:
        print("No se encontró el reporte ORGIQ más reciente en Descargas. Abortando.")
        sys.exit(1)

    ya_existe = set()
    for carpeta in (BORRADORES_DIR, DESPLEGADOS_DIR):
        if carpeta.exists():
            ya_existe.update(p.name for p in carpeta.iterdir() if p.suffix == ".docx")

    nuevos = []
    for f in filas:
        archivo_esperado = nombre_archivo(f["org"])
        if archivo_esperado in ya_existe:
            continue
        item = dict(f)
        item["archivo"] = archivo_esperado
        nuevos.append(item)

    # Prioriza las que tienen mas deals: son las unicas que pueden producir un
    # GBS real. Las que no tienen ninguno quedan al final (y con --limite,
    # normalmente ni entran).
    conteo = contar_deals_por_org()
    for item in nuevos:
        item["deals"] = conteo.get(item["org"], 0)
    nuevos.sort(key=lambda i: (-i["deals"], i["owner"], i["org"]))

    disponibles = len(nuevos)
    sin_deals = sum(1 for i in nuevos if i["deals"] == 0)
    recortados = 0
    if args.limite is not None and args.limite > 0 and disponibles > args.limite:
        recortados = disponibles - args.limite
        nuevos = nuevos[:args.limite]

    with open(PLAN_FILE, "w", encoding="utf-8") as f:
        json.dump({"items": nuevos}, f, ensure_ascii=False, indent=2)

    print(f"Organizaciones del equipo sin GBS Link (fuente: {orgiq_path.name}): {len(filas)}")
    print(f"Ya tienen borrador o GBS desplegado (se omiten): {len(filas) - disponibles}")
    print(f"Pendientes en total: {disponibles} (de esas, {sin_deals} no tienen ningún deal "
          f"asociado y probablemente no generen nada)")
    if recortados:
        print(f"Límite de esta corrida: {args.limite} — quedan {recortados} para la próxima vez.")
    print(f"Entran en esta corrida: {len(nuevos)}")
    for item in nuevos:
        print(f"  - {item['owner']}: {item['org']} ({item['deals']} deal(s))")


if __name__ == "__main__":
    main()
