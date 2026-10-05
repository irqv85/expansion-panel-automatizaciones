#!/usr/bin/env python3
"""
Paso 3: Extraer información de informes previos
Identifica account_id, extrae Sección 1 (esfuerzo comercial) y Sección 17 (veredicto)
"""

import json
import re
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "scripts"))
from paths import CHURN_BATCH_DIR, CHURN_INFORMES_DIR

def extract_section(md_content, section_num):
    """
    Extrae una sección del informe .md
    Las secciones están marcadas como "## Sección N:"
    """
    pattern = rf"##\s*Sección\s*{section_num}\s*:.*?\n(.*?)(?=##|$)"
    match = re.search(pattern, md_content, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None

def extract_field_from_table(md_content, field_name):
    """
    Extrae el valor de un campo desde una tabla Markdown
    Busca filas donde la primera columna contiene field_name
    Formato: | Campo | Valor |
    """
    # Buscar tabla que contiene el campo
    lines = md_content.split('\n')
    for i, line in enumerate(lines):
        # Buscar línea que contiene el campo (entre pipes |)
        if '|' in line and field_name.lower() in line.lower():
            # La siguiente línea podría ser el separador, o podría ser una fila
            # En tablas Markdown: | Campo | Valor |
            cells = [c.strip() for c in line.split('|')[1:-1]]  # Quitar pipes de inicio/fin
            if len(cells) >= 2:
                # Primera celda es el campo, segunda es el valor
                if field_name.lower() in cells[0].lower():
                    return cells[1] if len(cells) > 1 else None
    return None

def parse_effort(text):
    """
    Extrae el esfuerzo comercial del vendedor únicamente
    Si dice "NULO (vendedor) / RAZONABLE (Renewals)", devuelve solo "NULO"
    """
    if not text:
        return None

    # Si contiene paréntesis con roles, extraer solo la parte del vendedor
    vendor_match = re.search(r'(FUERTE|RAZONABLE|DÉBIL|NULO|INDETERMINADO)\s*\(\s*vendedor\s*\)', text, re.IGNORECASE)
    if vendor_match:
        return vendor_match.group(1)

    # Si contiene múltiples clasificaciones, tomar la primera
    effort_match = re.search(r'(FUERTE|RAZONABLE|DÉBIL|NULO|INDETERMINADO)', text, re.IGNORECASE)
    if effort_match:
        return effort_match.group(1)

    return None

def extract_verdict(md_content):
    """Extrae el veredicto de la Sección 17"""
    # Encontrar "17. VEREDICTO FINAL" (puede ser ## o * o nada)
    # seguido de una línea con el emoji + clasificación
    lines = md_content.split('\n')
    for i, line in enumerate(lines):
        if '17' in line and 'VEREDICTO' in line.upper() and 'FINAL' in line.upper():
            # Buscar en las próximas líneas la línea con el emoji
            for j in range(i+1, min(i+5, len(lines))):
                next_line = lines[j].strip()
                if '🟢' in next_line or '🟡' in next_line or '🟠' in next_line or '🔴' in next_line:
                    # Limpiar marcas de markdown
                    verdict = re.sub(r'\*\*|###', '', next_line).strip()
                    return verdict
    return None

def extract_effort_from_section1(md_content):
    """Extrae el esfuerzo comercial de la Sección 1"""
    # Buscar en toda la tabla de Sección 1 con pattern más flexible
    # La tabla tiene: | Esfuerzo comercial ... | VALOR |
    effort_match = re.search(r"\|\s*Esfuerzo\s+comercial[^|]*\|\s*([^|]+)\|", md_content, re.IGNORECASE)
    if effort_match:
        text = effort_match.group(1).strip()
        return parse_effort(text)
    return None

def load_and_parse_report(md_file_path):
    """Carga un informe .md y extrae la información necesaria"""
    with open(md_file_path, 'r', encoding='utf-8') as f:
        md_content = f.read()

    # Extraer campos
    effort = extract_effort_from_section1(md_content)
    verdict = extract_verdict(md_content)

    # Extraer account_id del nombre del archivo
    filename = md_file_path.stem
    account_id_match = re.search(r'(3x\d+)', filename)
    account_id = account_id_match.group(1) if account_id_match else None

    # Extraer fecha del informe (último componente del nombre, ej. 2026-08-26)
    date_match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)
    report_date = date_match.group(1) if date_match else None

    return {
        'account_id': account_id,
        'report_date': report_date,
        'md_path': str(md_file_path),
        'html_path': str(md_file_path.with_suffix('.html')),
        'effort': effort,
        'verdict': verdict,
        'full_content': md_content
    }

def slug_organizacion(nombre):
    """Mismo slug que usa la skill analiza-churn para nombrar el informe:
    minusculas y todo lo que no sea alfanumerico convertido en guion."""
    return re.sub(r"[^a-z0-9]+", "-", nombre.lower()).strip("-")


def buscar_informe(organization):
    """Devuelve el .md mas reciente de esa organizacion, o None.

    Se busca por slug del nombre y no por account_id porque el account_id no
    siempre esta resuelto: paso2 lo deja en null (no tiene MCP de vTiger) y hay
    informes viejos guardados sin el prefijo 3x, como
    cima-group-sa-de-cv_9485060_2026-08-24.md. El slug si esta siempre.
    """
    patron = slug_organizacion(organization) + "_*.md"
    md_files = list(CHURN_INFORMES_DIR.glob(patron))
    if not md_files:
        return None
    return sorted(md_files, key=lambda q: q.stat().st_mtime, reverse=True)[0]


def main():
    print("=== PASO 3: Extraer informacion de informes previos ===")
    print()

    # Hasta el 1-oct-2026 este paso llevaba un accounts_map escrito a mano con
    # 3 cuentas de agosto, asi que ignoraba cualquier cuenta nueva por muchos
    # informes que hubiera en la carpeta. Ahora la lista sale de paso2.
    entrada = CHURN_BATCH_DIR / "cuentas_para_analizar.json"
    if not entrada.exists():
        print("ERROR: falta " + str(entrada) + ". Corre antes paso1 y paso2.")
        sys.exit(1)
    with open(entrada, encoding="utf-8") as f:
        cuentas = json.load(f)

    # Una organizacion puede traer varios assets en churn (CIMA trae 2) y el
    # informe forense es uno por cuenta, no por asset.
    unicas = {}
    for c in cuentas:
        unicas.setdefault(c["org_id"], c)

    reports_data = []
    pendientes = []

    for org_id, org_info in unicas.items():
        organization = org_info["organization"]
        print("Buscando informe para " + organization + " (" + org_id + ")...")

        md_file = buscar_informe(organization)
        if not md_file:
            print("  SIN INFORME PREVIO: hay que correrle el analisis forense")
            print()
            pendientes.append(org_info)
            continue

        print("  Encontrado: " + md_file.name)
        report_info = load_and_parse_report(md_file)
        report_info.update({
            "org_id": org_id,
            "organization": organization,
            "assigned_to": org_info.get("assigned_to"),
        })
        # El account_id fiable es el del nombre del informe; el de paso2 viene
        # en null mientras no pase por vTiger.
        if not report_info.get("account_id"):
            report_info["account_id"] = org_info.get("account_id")

        html_file = md_file.with_suffix(".html")
        if html_file.exists():
            print("    HTML: " + html_file.name)
        else:
            print("    ADVERTENCIA: no existe " + html_file.name)

        print("    Esfuerzo comercial: " + (report_info["effort"] or "NO ENCONTRADO"))
        print("    Veredicto: " + (report_info["verdict"] or "NO ENCONTRADO"))
        print()
        reports_data.append(report_info)

    output_file = CHURN_BATCH_DIR / "informes_extraidos.json"

    reports_clean = []
    for r in reports_data:
        reports_clean.append({
            "org_id": r["org_id"],
            "organization": r["organization"],
            "account_id": r["account_id"],
            "assigned_to": r["assigned_to"],
            "report_date": r["report_date"],
            "md_path": r["md_path"],
            "html_path": r["html_path"],
            "effort": r["effort"],
            "verdict": r["verdict"],
        })

    # Windows escribe en ANSI por defecto: sin encoding explicito esto reventaba
    # con UnicodeEncodeError en el emoji del veredicto y dejaba el JSON a medias,
    # con lo que el paso 5 se quedaba sin insumo y no salia el Excel (1-oct-2026).
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(reports_clean, f, indent=2, ensure_ascii=False)

    pendientes_file = CHURN_BATCH_DIR / "cuentas_pendientes.json"
    with open(pendientes_file, "w", encoding="utf-8") as f:
        json.dump(pendientes, f, indent=2, ensure_ascii=False)

    print()
    print("Informacion extraida guardada en: " + str(output_file))
    print("Total cuentas reutilizadas: " + str(len(reports_data)))
    if pendientes:
        print("Cuentas SIN informe previo (" + str(len(pendientes)) +
              "), van al analisis forense:")
        for c in pendientes:
            print("  - " + c["organization"] + " (" + c["org_id"] + ")")
        print("Guardadas en: " + str(pendientes_file))
    else:
        print("Todas las cuentas tenian informe previo.")


if __name__ == '__main__':
    main()
