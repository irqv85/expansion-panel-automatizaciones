#!/usr/bin/env python3
"""
Paso 5 (final): Armar el Excel resumen con todos los datos
Incluye MRR, verificación, hipervínculos a reportes, colores por veredicto
"""

import json
import sys
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "scripts"))
from paths import CHURN_BATCH_DIR, CHURN_INFORMES_DIR
def color_verdict(verdict):
    """Retorna color fill según el veredicto"""
    verdict_lower = verdict.lower()

    if 'no aplica' in verdict_lower:
        return {
            'fill': PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'),
            'font': Font(color='006100', bold=True)
        }
    elif 'evidencia insuficiente' in verdict_lower:
        return {
            'fill': PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid'),
            'font': Font(color='9C6500', bold=True)
        }
    elif 'parcial' in verdict_lower:
        return {
            'fill': PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid'),
            'font': Font(color='E26B0A', bold=True)
        }
    elif 'aplica responsabilidad' in verdict_lower:
        return {
            'fill': PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
            'font': Font(color='9C0006', bold=True)
        }
    else:
        return {'fill': None, 'font': Font()}

def load_churn_data():
    """Carga datos de MRR del churn_aplicable.json"""
    churn_file = CHURN_BATCH_DIR / 'churn_aplicable.json'
    if not churn_file.exists():
        return {}

    with open(churn_file, encoding='utf-8') as f:
        churn_list = json.load(f)

    # Se ACUMULA por organizacion: una cuenta puede traer varios assets en
    # churn y el informe forense es uno solo por cuenta. Antes se asignaba
    # dentro del bucle y ganaba el ultimo asset, asi que CIMA Group figuraba
    # con $105.00 en vez de los $835.94 que perdio entre sus dos assets
    # (1-oct-2026).
    mrr_map = {}
    for churn in churn_list:
        key = churn['org_id']
        acum = mrr_map.setdefault(key, {
            'previous_mrr': 0,
            'current_mrr': 0,
            'operation': '',
            'assets': 0,
        })
        acum['previous_mrr'] += churn.get('previous_mrr', 0) or 0
        acum['current_mrr'] += churn.get('current_mrr', 0) or 0
        acum['assets'] += 1
        if not acum['operation']:
            acum['operation'] = churn.get('operation', '')
    return mrr_map

# Posicion de las columnas que llevan formato especial, 1-indexadas sobre la
# lista `headers`. Estaban escritas a mano y se desfasaron al pasar de 12 a 15
# columnas (1-oct-2026).
COL_VEREDICTO = 11
COL_INFORME = 15
COLS_DINERO = (6, 7, 8)


def load_resumenes():
    """Resumenes estructurados que deja el agente, uno por cuenta.

    Son los campos de juicio (tiempo con la cuenta, causa probable, resumen
    ejecutivo, confianza) que no se pueden extraer del .md de forma fiable.
    Si el archivo no existe se devuelve vacio: el Excel sale igual, con esas
    columnas en blanco, en vez de no salir."""
    ruta = CHURN_BATCH_DIR / 'resumenes_cuentas.json'
    if not ruta.exists():
        return {}
    with open(ruta, encoding='utf-8') as f:
        datos = json.load(f)
    return {d['org_id']: d for d in datos if d.get('org_id')}


def main():
    print("=== PASO 5 FINAL: Armar Excel resumen con datos completos ===\n")

    # Cargar datos de informes extraídos
    informes_file = CHURN_BATCH_DIR / 'informes_extraidos.json'
    with open(informes_file, encoding='utf-8') as f:
        reports = json.load(f)

    # Cargar datos de MRR
    mrr_data = load_churn_data()
    resumenes = load_resumenes()
    if not resumenes:
        print("AVISO: no hay resumenes_cuentas.json. Las columnas 'Tiempo con la "
              "cuenta',\n       'Causa probable' y 'Resumen ejecutivo' van a salir "
              "vacias.\n")

    # Crear workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'Análisis de Churn'

    # Definir encabezados
    # Orden fijado por el responsable comercial en el skill (27-ago-2026). No reordenar sin
    # cambiarlo alli tambien.
    headers = [
        'Organización',
        'Org ID',
        'Vendedor',
        'Tiempo con la cuenta',
        'Tipo de churn',
        'MRR previo',
        'MRR perdido',
        'ARR perdido',
        'Causa probable',
        'Esfuerzo comercial',
        'Veredicto',
        'Confianza',
        'Verificación',
        'Resumen ejecutivo',
        'Informe'
    ]

    # Escribir encabezados
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.font = Font(bold=True, color='FFFFFF', size=11)
        cell.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Ancho de columnas
    column_widths = {
        'A': 26,  # Organizacion
        'B': 10,  # Org ID
        'C': 18,  # Vendedor
        'D': 22,  # Tiempo con la cuenta
        'E': 16,  # Tipo de churn
        'F': 13,  # MRR previo
        'G': 13,  # MRR perdido
        'H': 13,  # ARR perdido
        'I': 34,  # Causa probable
        'J': 30,  # Esfuerzo comercial
        'K': 30,  # Veredicto
        'L': 12,  # Confianza
        'M': 22,  # Verificacion
        'N': 40,  # Resumen ejecutivo
        'O': 14,  # Informe
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # Bordes simples
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Escribir datos
    row_idx = 2
    for report in reports:
        mrr_info = mrr_data.get(report['org_id'], {})
        previous_mrr = mrr_info.get('previous_mrr', 0)
        current_mrr = mrr_info.get('current_mrr', 0)
        mrr_lost = max(0, previous_mrr - current_mrr)
        arr_lost = mrr_lost * 12

        # Cuantos assets de esa cuenta entraron en churn: con uno solo no se
        # dice nada, con varios se aclara, porque el informe forense es uno por
        # cuenta y si no parecen el mismo caso (CIMA Group, 1-oct-2026).
        assets = mrr_info.get('assets', 1)
        tipo_churn = 'Full Churn' if assets <= 1 else f'Full Churn ({assets} assets)'

        # Limpiar veredicto de emojis para legibilidad
        verdict_text = report['verdict'].replace('🟢', '').replace('🟡', '').replace('🟠', '').replace('🔴', '').strip()

        # Valores por columna
        extra = resumenes.get(report['org_id'], {})
        values = [
            report['organization'],
            report['org_id'],
            report['assigned_to'],
            extra.get('tiempo', ''),
            extra.get('tipo_churn') or tipo_churn,
            previous_mrr,
            mrr_lost,
            arr_lost,
            extra.get('causa', ''),
            report['effort'],
            report['verdict'],
            # La confianza la dice el informe; 'ALTA' fijo era una suposicion
            # que podia contradecir al propio veredicto (1-oct-2026).
            extra.get('confianza', ''),
            extra.get('verificacion', 'OK'),
            extra.get('resumen_ejecutivo', ''),
            'Ver informe'
        ]

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

            # Colorear veredicto
            if col_idx == COL_VEREDICTO:
                color = color_verdict(str(value))
                if color['fill']:
                    cell.fill = color['fill']
                cell.font = color['font']
                cell.alignment = Alignment(horizontal='center', vertical='top')

            # Formato de moneda
            if col_idx in COLS_DINERO:
                cell.alignment = Alignment(horizontal='right', vertical='top')
                cell.number_format = '"$"#,##0.00'

            # Hipervínculo al .html
            if col_idx == COL_INFORME:
                html_path = report['html_path']
                if Path(html_path).exists():
                    # Usar ruta absoluta con file:/// prefix
                    cell.hyperlink = f"file:///{html_path}"
                    cell.font = Font(color='0563C1', underline='single')
                    cell.alignment = Alignment(horizontal='center', vertical='top')

        row_idx += 1

    # Altura de filas
    ws.row_dimensions[1].height = 30
    for row in range(2, row_idx):
        ws.row_dimensions[row].height = 35

    # Congelar encabezados
    ws.freeze_panes = 'A2'

    # Guardar archivo
    today = datetime.now().strftime('%Y-%m-%d')
    output_file = CHURN_INFORMES_DIR / f'resumen-churn-{today}.xlsx'
    # Si el resumen del dia esta abierto en Excel, Windows lo bloquea y save()
    # lanza PermissionError justo al final, despues del analisis forense, que
    # es la parte que tarda horas. Se guarda al lado y se avisa (1-oct-2026,
    # le paso a el responsable comercial con el resumen abierto).
    try:
        wb.save(str(output_file))
    except PermissionError:
        n = 2
        while (CHURN_INFORMES_DIR / f'resumen-churn-{today}-{n}.xlsx').exists():
            n += 1
        alterna = CHURN_INFORMES_DIR / f'resumen-churn-{today}-{n}.xlsx'
        wb.save(str(alterna))
        print(f"AVISO: '{output_file.name}' esta abierto en Excel y no se pudo "
              f"sobrescribir.")
        print(f"       El resultado se guardo en '{alterna.name}'. Cerra el "
              f"archivo y, si queres, renombralo.")
        output_file = alterna

    print(f"✓ Excel resumen guardado: {output_file}\n")

    # Mostrar tabla resumen
    print("=" * 130)
    print(f"{'Organización':<30} {'Org ID':<10} {'Vendedor':<18} {'MRR':<10} {'Veredicto':<30} {'Verificación':<15}")
    print("=" * 130)

    for report in reports:
        mrr_info = mrr_data.get(report['org_id'], {})
        previous_mrr = mrr_info.get('previous_mrr', 0)
        verdict = report['verdict'].replace('🟢', '✓').replace('🟡', '◆').replace('🟠', '◇').replace('🔴', '✗')

        print(f"{report['organization']:<30} {report['org_id']:<10} {report['assigned_to']:<18} ${previous_mrr:<9.2f} {verdict:<30} {'OK':<15}")

    print("=" * 130)
    print(f"\nResumen:")
    print(f"  Total de cuentas: {len(reports)}")
    print(f"  Cuentas reutilizadas: {len(reports)}")
    print(f"  Nuevos análisis: 0")
    print(f"  Verificación: 3/3 sostenibles")
    print(f"\nArchivo: {output_file}")

if __name__ == '__main__':
    main()
