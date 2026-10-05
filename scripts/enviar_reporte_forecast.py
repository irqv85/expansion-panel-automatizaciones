"""
Envia por Outlook el ultimo reporte de Forecast generado (segun
ultima_generacion_forecast.json) a la lista de destinatarios configurada en
"Forecast - Destinatarios.xlsx" (mismo carpeta que este script).

Ese Excel lo edita el responsable comercial directamente, sin tocar codigo: una fila por
persona, con su email y si va en "Para" o "CC". Si el archivo no existe
todavia, se crea un molde EN BLANCO (sin destinatarios precargados -- este
reporte va para Freshworks, no para el equipo interno, asi que la lista la
arma el responsable comercial, no el script).

Uso:
  python enviar_reporte_forecast.py             Envia el ultimo reporte generado
  python enviar_reporte_forecast.py --dry-run   Solo muestra que haria
"""
import argparse
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

from paths import FORECAST_DIR

SCRIPT_DIR = Path(__file__).resolve().parent
MANIFEST_FILE = SCRIPT_DIR / "ultima_generacion_forecast.json"
DESTINATARIOS_FILE = SCRIPT_DIR / "Forecast - Destinatarios.xlsx"

def crear_destinatarios_defecto():
    """Molde en blanco: sin nombres precargados. Este reporte es para
    Freshworks, no para el equipo -- el responsable comercial decide quien va, no el script."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Destinatarios"
    ws["A1"] = "Nombre"
    ws["B1"] = "Email"
    ws["C1"] = "Tipo (Para / CC)"
    for c in ("A1", "B1", "C1"):
        ws[c].font = Font(bold=True)
    ws["A2"] = "Agrega aqui a quien le corresponda recibir el Forecast (ej. tu contacto en Freshworks). Una fila por persona. Tipo acepta 'Para' o 'CC' (si no pones nada, se manda 'Para')."
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 18
    wb.save(DESTINATARIOS_FILE)


def leer_destinatarios():
    wb = openpyxl.load_workbook(DESTINATARIOS_FILE, data_only=True)
    ws = wb[wb.sheetnames[0]]
    para, cc = [], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        email = str(row[1]).strip()
        if "@" not in email:
            continue
        tipo = str(row[2] or "").strip().lower() if len(row) > 2 else ""
        if tipo == "cc":
            cc.append(email)
        else:
            para.append(email)
    wb.close()
    return para, cc


def main():
    parser = argparse.ArgumentParser(description="Envia el ultimo reporte de Forecast generado.")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    if not MANIFEST_FILE.exists():
        print("Todavía no hay un Forecast generado. Corre generar_reporte_forecast.py primero.")
        return

    with open(MANIFEST_FILE, "r", encoding="utf-8") as f:
        manifiesto = json.load(f)
    reporte_path = FORECAST_DIR / manifiesto["archivo"]
    if not reporte_path.exists():
        print(f"El reporte del manifiesto no existe: {reporte_path}")
        return

    if not DESTINATARIOS_FILE.exists():
        crear_destinatarios_defecto()
        print(f'Se creó "{DESTINATARIOS_FILE.name}" con el equipo por defecto. '
              "Revísalo/ajústalo y vuelve a correr este script para enviar.")
        return

    para, cc = leer_destinatarios()
    if not para and not cc:
        print(f'"{DESTINATARIOS_FILE.name}" no tiene ningún email válido. Agrega al menos uno y reintenta.')
        return

    mes_label = manifiesto.get("anio_mes_objetivo", "")
    subject = f"Forecast Mid/SMB - {mes_label}"
    body = (
        "Hola equipo,\n\n"
        "Les comparto el forecast (Upside y Commit) del período correspondiente.\n\n"
        "Saludos."
    )

    if args.dry_run:
        print(f"[DRY-RUN] Para={'; '.join(para)} CC={'; '.join(cc)} "
              f"Subject='{subject}' Adjunto={reporte_path.name}")
        return

    # Supervisado a proposito: este reporte va a Freshworks y los
    # destinatarios los elige el responsable comercial caso por caso, asi que se abre la ventana
    # de redaccion y el clic final en "Enviar" queda en sus manos.
    import mail_helper
    mail_helper.display_mail(
        to=para, cc=cc or None, subject=subject, body=body, attachment=reporte_path,
    )
    print(f"Outlook abierto con '{reporte_path.name}' para: {'; '.join(para)}"
          + (f" (CC: {'; '.join(cc)})" if cc else "")
          + " -- confirma el envío dando clic en Enviar dentro de Outlook.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        traceback.print_exc()
        sys.exit(1)
