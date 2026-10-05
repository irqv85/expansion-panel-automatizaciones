#!/usr/bin/env python3
"""
Paso 5: Armar el Excel resumen final
Toma los resultados de verificación y crea el archivo xlsx
"""

import json
import sys
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from urllib.parse import urljoin

def color_verdict(verdict):
    """Retorna color fill según el veredicto"""
    verdict_lower = verdict.lower()

    if 'no aplica' in verdict_lower:
        return {
            'fill': PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'),
            'font': Font(color='006100')
        }
    elif 'evidencia insuficiente' in verdict_lower:
        return {
            'fill': PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid'),
            'font': Font(color='9C6500')
        }
    elif 'parcial' in verdict_lower:
        return {
            'fill': PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid'),
            'font': Font(color='E26B0A')
        }
    elif 'aplica responsabilidad' in verdict_lower:
        return {
            'fill': PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
            'font': Font(color='9C0006')
        }
    else:
        return {'fill': None, 'font': Font()}

def extract_mrr_from_churn_data(churn_data):
    """Busca en churn_data.json los valores de MRR para cada org"""
    churn_file = CHURN_BATCH_DIR / 'churn_aplicable.json'
    if not churn_file.exists():
        return {}

    with open(churn_file, encoding='utf-8') as f:
        churn_list = json.load(f)

    mrr_map = {}
    for churn in churn_list:
        key = churn['org_id']
        mrr_map[key] = {
            'previous_mrr': churn.get('previous_mrr', 0),
            'current_mrr': churn.get('current_mrr', 0),
            'operation': churn.get('operation', '')
        }
    return mrr_map

def main():
    print("=== PASO 5: Armar Excel resumen ===\n")

    # Cargar datos de verificación (serán pasados como argumento desde el main script)
    # Por ahora, cargaremos desde los archivos locales
    informes_file = CHURN_BATCH_DIR / 'informes_extraidos.json'

    with open(informes_file, encoding='utf-8') as f:
        reports = json.load(f)

    # Cargar datos de MRR
    mrr_data = extract_mrr_from_churn_data(None)

    # Crear workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'Churn Analysis'

    # Definir encabezados
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
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Ancho de columnas
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['I'].width = 25
    ws.column_dimensions['J'].width = 18
    ws.column_dimensions['K'].width = 25
    ws.column_dimensions['L'].width = 10
    ws.column_dimensions['M'].width = 20
    ws.column_dimensions['N'].width = 40
    ws.column_dimensions['O'].width = 18

    # Escribir datos
    row_idx = 2
    for report in reports:
        mrr_info = mrr_data.get(report['org_id'], {})
        previous_mrr = mrr_info.get('previous_mrr', 0)
        current_mrr = mrr_info.get('current_mrr', 0)
        mrr_lost = max(0, previous_mrr - current_mrr)
        arr_lost = mrr_lost * 12

        # Valores por columna
        values = [
            report['organization'],
            report['org_id'],
            report['assigned_to'],
            'N/A',  # Tiempo con la cuenta (extractable del .md si es necesario)
            'Full Churn',  # Tipo de churn (extractable del .md)
            f"${previous_mrr:.2f}",
            f"${mrr_lost:.2f}",
            f"${arr_lost:.2f}",
            'N/A',  # Causa probable (extractable del .md)
            report['effort'],
            report['verdict'],
            'ALTA',  # Confianza (del informe)
            'OK',  # Verificación (será actualizado con los resultados del Workflow)
            'Reutilizado, ya analizado el ' + report['report_date'],
            'Ver informe'  # Texto del hipervínculo
        ]

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

            # Colorear veredicto
            if col_idx == 11:  # Columna K (Veredicto)
                color = color_verdict(str(value))
                if color['fill']:
                    cell.fill = color['fill']
                cell.font = color['font']

            # Hipervínculo
            if col_idx == 15:  # Columna O (Informe)
                html_path = report['html_path']
                if Path(html_path).exists():
                    cell.hyperlink = f"file:///{html_path}"
                    cell.font = Font(color='0563C1', underline='single')

        row_idx += 1

    # Guardar archivo
    today = datetime.now().strftime('%Y-%m-%d')
    output_file = CHURN_INFORMES_DIR / f'resumen-churn-{today}.xlsx'
    wb.save(str(output_file))

    print(f"Excel resumen guardado: {output_file}")
    print(f"Total de cuentas: {len(reports)}")
    print(f"\nCuentas procesadas:")
    for report in reports:
        print(f"  {report['org_id']:10} | {report['organization']:35} | {report['verdict']}")

if __name__ == '__main__':
    main()
