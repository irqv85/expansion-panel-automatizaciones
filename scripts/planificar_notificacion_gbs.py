"""
Calcula a quien le falta notificar del ultimo chequeo de GBS Deal Checker:
compara los deals con problema del Excel de HOY (GBS Deal Checker/informes/
gbs-checker-YYYY-MM-DD.xlsx) contra el registro de ya notificados
(state_gbs_notificaciones.json) y escribe el plan resultante en
plan_notificacion_gbs.json. No escribe nada en vTiger -- eso lo hace Claude
Code despues, siguiendo el skill "notificar-gbs-deal-checker" con este plan
ya calculado (mismo patron "Calcular -> Enviar" que Sofia/Calidad CRM).

Un deal que se arregla desaparece del Excel de "problema" y por lo tanto
tambien se poda del registro de notificados -- si se rompe de nuevo mas
adelante, se vuelve a notificar como si fuera nuevo.

Ejecutar con: python planificar_notificacion_gbs.py
"""
import json
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

from paths import GBS_CHECKER_INFORMES_DIR as INFORMES_DIR

SCRIPT_DIR = Path(__file__).resolve().parent
STATE_FILE = SCRIPT_DIR / "state_gbs_notificaciones.json"
PLAN_FILE = SCRIPT_DIR / "plan_notificacion_gbs.json"


def deal_key(owner, deal, org):
    return f"{owner}||{deal}||{org}"


def cargar_problema_hoy():
    hoy = datetime.now().strftime("%Y-%m-%d")
    archivo = INFORMES_DIR / f"gbs-checker-{hoy}.xlsx"
    if not archivo.exists():
        return None, archivo

    wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    ws = wb.active
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = {h: i for i, h in enumerate(header)}
    filas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue
        filas.append({
            "owner": row[idx["Representante"]],
            "deal": row[idx["Deal"]],
            "org": row[idx["Organización"]],
            "etapa": row[idx["Etapa"]],
            "estado": row[idx["Estado GBS Link"]],
            "link": row[idx["Link"]],
        })
    wb.close()
    return filas, archivo


# Dias que deben pasar antes de volver a insistirle a un vendedor sobre el
# MISMO deal. Antes del 24-sep-2026 no se insistia nunca: se avisaba una vez y
# el deal quedaba marcado para siempre, asi que quien ignorara la mencion se
# quedaba tranquilo. Ese dia habia 8 deals avisados el 28-ago, 27 dias antes,
# todavia sin GBS. 7 dias deja margen para corregirlo sin convertirlo en spam
# si se pulsa el boton varios dias seguidos.
DIAS_PARA_INSISTIR = 7


def _normalizar_registro(valor):
    """El registro guardaba solo la fecha del aviso ("2026-08-28"). Ahora
    guarda tambien cuantas veces se aviso, para que el tono del comentario
    pueda escalar. Se acepta el formato viejo y se migra al leerlo."""
    if isinstance(valor, str):
        return {"primera": valor, "ultima": valor, "avisos": 1}
    if isinstance(valor, dict):
        ultima = valor.get("ultima") or valor.get("primera")
        return {
            "primera": valor.get("primera") or ultima,
            "ultima": ultima,
            "avisos": int(valor.get("avisos", 1) or 1),
        }
    return None


def cargar_estado():
    if STATE_FILE.exists():
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                crudos = data.get("notificados", {}) or {}
                data["notificados"] = {
                    k: r for k, r in ((k, _normalizar_registro(v)) for k, v in crudos.items())
                    if r
                }
                return data
        except Exception:
            pass
    return {"notificados": {}}


def guardar_estado(estado):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


def main():
    filas, archivo = cargar_problema_hoy()
    if filas is None:
        print(f"No hay un chequeo de GBS de hoy ({archivo.name} no existe). "
              "Corre 'Chequear GBS' primero.")
        sys.exit(1)

    estado = cargar_estado()
    notificados = estado["notificados"]

    claves_hoy = {deal_key(f["owner"], f["deal"], f["org"]) for f in filas}
    podados = [k for k in list(notificados) if k not in claves_hoy]
    for k in podados:
        del notificados[k]

    hoy_dt = datetime.now().date()
    items = []
    esperando = 0
    for f in filas:
        registro = notificados.get(deal_key(f["owner"], f["deal"], f["org"]))
        if not registro:
            items.append(dict(f, tipo="primera", avisos_previos=0))
            continue
        try:
            dias = (hoy_dt - datetime.strptime(registro["ultima"], "%Y-%m-%d").date()).days
        except Exception:
            dias = DIAS_PARA_INSISTIR  # fecha ilegible: mejor insistir que callar
        if dias < DIAS_PARA_INSISTIR:
            esperando += 1
            continue
        items.append(dict(
            f, tipo="recordatorio", avisos_previos=registro["avisos"],
            primera_notificacion=registro["primera"],
            ultima_notificacion=registro["ultima"], dias_sin_corregir=dias,
        ))

    nuevos = [i for i in items if i["tipo"] == "primera"]
    recordatorios = [i for i in items if i["tipo"] == "recordatorio"]

    with open(PLAN_FILE, "w", encoding="utf-8") as f:
        json.dump({"items": items}, f, ensure_ascii=False, indent=2)
    guardar_estado(estado)

    print(f"Deals con problema hoy ({archivo.name}): {len(filas)}")
    print(f"Por notificar por primera vez: {len(nuevos)}")
    print(f"Recordatorios (avisados hace {DIAS_PARA_INSISTIR}+ dias y sin corregir): "
          f"{len(recordatorios)}")
    if esperando:
        print(f"Avisados hace menos de {DIAS_PARA_INSISTIR} dias (se les da tiempo): {esperando}")
    if podados:
        print(f"Se destildaron {len(podados)} deal(es) del registro de notificados "
              "(ya no aparecen con problema, se corrigieron).")
    for item in items:
        if item["tipo"] == "primera":
            print(f"  - [nuevo] {item['owner']}: {item['deal']} ({item['org']}) — {item['estado']}")
        else:
            print(f"  - [RECORDATORIO x{item['avisos_previos'] + 1}, "
                  f"{item['dias_sin_corregir']} dias sin corregir] {item['owner']}: "
                  f"{item['deal']} ({item['org']}) — {item['estado']}")


if __name__ == "__main__":
    main()
