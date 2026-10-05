#!/usr/bin/env python3
"""
Paso 2: Resolver account_id en vTiger y verificar informes previos
"""

import json
from pathlib import Path
import sys
import re
from datetime import datetime

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "scripts"))
from paths import CHURN_BATCH_DIR, CHURN_INFORMES_DIR

# Importar MCP (tools de vTiger disponibles en el entorno)
try:
    from mcp import vtiger_query, vtiger_authenticate
except ImportError:
    print("NOTA: Usaremos bash/curl para conectar a vTiger, no la importación MCP")
    vtiger_query = None

# Para este script, vamos a usar un enfoque híbrido:
# El usuario debe ejecutar consultas a vTiger manualmente o usamos una herramienta disponible

def load_churn_aplicable():
    """Carga la lista de churn aplicable del paso anterior"""
    json_file = CHURN_BATCH_DIR / 'churn_aplicable.json'
    with open(json_file, encoding='utf-8') as f:
        return json.load(f)

def find_existing_reports(account_id_or_name):
    """
    Busca informes previos para una cuenta en el directorio informes/
    Busca archivos con patrón: <slug>_<account_id>_*.md
    """
    informes_dir = CHURN_INFORMES_DIR
    if not informes_dir.exists():
        return []

    # account_id tiene formato 3x<N>, ej 3x1234
    # Buscar archivos que contengan _3x en el nombre (formato estándar del análisis)
    pattern = re.compile(r'_3x\d+_')

    reports = []
    for md_file in informes_dir.glob('*.md'):
        if pattern.search(md_file.name):
            reports.append(md_file)

    return reports

def find_related_html(md_file):
    """Busca el .html correspondiente a un .md"""
    html_file = md_file.with_suffix('.html')
    if html_file.exists():
        return html_file
    return None

def main():
    print("=== PASO 2: Resolver account_id y verificar informes previos ===\n")

    churn_data = load_churn_aplicable()
    print(f"Cuentas con churn aplicable: {len(churn_data)}\n")

    # Aquí necesitamos conectar a vTiger para resolver los account_id
    # El skill indica:
    # - username: el de la variable VTIGER_USERNAME del .env
    # - accessKey: el de la variable VTIGER_ACCESS_KEY del .env
    # - Query: SELECT id, accountname, account_no FROM Accounts WHERE account_no='<org_id>'

    print("INSTRUCCIÓN: Para resolver los account_id en vTiger, se necesita consultar cada cuenta.")
    print("Como el MCP de vTiger no está disponible en este script standalone,")
    print("vamos a guardar los datos que necesitamos y continuar con el análisis.\n")

    # Por ahora, guardamos la estructura que necesitaremos pasar al Workflow
    # El Workflow es quien tendrá acceso a los tools MCP

    result = []
    for org in churn_data:
        account_info = {
            'org_id': org['org_id'],
            'organization': org['organization'],
            'asset_name': org['asset_name'],
            'assigned_to': org['assigned_to'],
            'tier': org['tier'],
            'previous_mrr': org['previous_mrr'],
            'current_mrr': org['current_mrr'],
            'operation': org['operation'],
            'active_mrr': org['active_mrr'],
            'account_id': None,  # Será resuelto por el Workflow
            'has_existing_report': None,  # Será verificado después
        }
        result.append(account_info)

    # Guardar para el siguiente paso
    output_file = CHURN_BATCH_DIR / 'cuentas_para_analizar.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    print(f"Información de cuentas guardada en: {output_file}")
    print("\nCuentas a procesar:")
    print("-" * 100)
    for org in result:
        print(f"{org['org_id']:10} | {org['organization']:35} | {org['assigned_to']}")

    print(f"\nProxImo paso: El Workflow resolverá los account_id en vTiger")
    print("y verificará si existen informes previos.")

if __name__ == '__main__':
    main()
