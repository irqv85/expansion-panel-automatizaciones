"""Asignar el producto a los farmings que no lo tienen.

Dos pasos, como el boton: LISTAR (solo lectura) y AÑADIR (escribe en vTiger).
Esta logica no depende de la pantalla, asi se puede probar sin tocar el CRM.

Por que el producto no se deduce solo del nombre del farming. Los farmings sin
producto de esta cuenta se llaman "2026 - CSM - <cliente>" o "2025 - ITSM -
<cliente>": CSM e ITSM son categorias, no productos, y un mismo "CSM" aparece con
Freshdesk en unos casos y con Freshdesk Omni 2026 en otros. Adivinar ahi seria
escribir datos falsos en el CRM. Solo se sugiere cuando el nombre trae el
producto EXACTO, como "2026 - NinjaOne - <cliente>"; en lo demas decide la
persona.
"""
import html
import re

import vtiger_api as v

MODULO = "vtcmfarming"
CAMPO = "cf_vtcmfarming_product"

_CAMPOS = ("id, fld_vtcmfarmingname, cf_vtcmfarming_farmingstage, "
           "cf_vtcmfarming_product, assigned_user_id")


def productos_validos(sesion):
    """Valores del picklist, tal como los guarda vTiger.

    Se devuelven pares (valor_crudo, texto_para_mostrar). El crudo es lo que hay
    que escribir de vuelta: vTiger entrega algunos con entidades HTML
    ("AT&amp;T") y caracteres invisibles, y reescribirlos distintos fallaria o
    crearia un valor nuevo. El texto de pantalla se limpia solo para leerlo."""
    desc = v.describir(sesion, MODULO)
    campo = next(f for f in desc["fields"] if f["name"] == CAMPO)
    pares = []
    for x in campo["type"]["picklistValues"]:
        crudo = x["value"]
        mostrar = html.unescape(crudo).replace("​", "").strip()
        pares.append((crudo, mostrar))
    return pares


def sugerir(nombre, pares):
    """Producto sugerido a partir del nombre del farming, o None.

    Exige que el nombre contenga el producto completo y como palabra suelta:
    asi "Halo" no se confunde con "Halogen". Si dos productos encajan, gana el
    mas largo ("ServiceNow - ITSM" antes que "Servicenow")."""
    texto = nombre.lower()
    mejor = None
    for crudo, mostrar in pares:
        clave = mostrar.lower()
        if len(clave) < 3:
            continue
        if re.search(r"(?<![a-z0-9])" + re.escape(clave) + r"(?![a-z0-9])", texto):
            if mejor is None or len(clave) > len(mejor[1]):
                mejor = (crudo, clave)
    return mejor[0] if mejor else None


def listar_sin_producto(sesion, mi_id, pares):
    """Farmings asignados a `mi_id` que no tienen producto. SOLO LECTURA."""
    filas = v.consultar_todo(
        sesion,
        f"SELECT {_CAMPOS} FROM {MODULO} WHERE assigned_user_id = '{mi_id}'",
    )
    salida = []
    for f in filas:
        if f.get(CAMPO):
            continue
        nombre = f.get("fld_vtcmfarmingname") or ""
        salida.append({
            "id": f["id"],
            "nombre": nombre,
            "etapa": f.get("cf_vtcmfarming_farmingstage") or "",
            "sugerido": sugerir(nombre, pares),
        })
    salida.sort(key=lambda r: r["nombre"].lower())
    return salida, len(filas)


def aplicar(sesion, asignaciones, pares, escribir=v.revisar):
    """Escribe el producto elegido en cada farming. ESCRIBE EN vTiger.

    `asignaciones` es una lista de (id, nombre, producto_crudo). Cada una se
    intenta por separado: si una falla, las demas siguen, y al final se informa
    cual salio y cual no. Parar a la primera dejaria la mitad hecha sin decirlo.

    Antes de escribir se comprueba que el producto este en la lista valida: vTiger
    rechazaria un valor ajeno, pero es mejor no mandarlo."""
    validos = {crudo for crudo, _ in pares}
    ok, fallos = [], []
    for id_registro, nombre, producto in asignaciones:
        if producto not in validos:
            fallos.append((nombre, f"«{producto}» no esta en la lista de productos de vTiger"))
            continue
        try:
            escribir(sesion, id_registro, {CAMPO: producto})
            ok.append((nombre, producto))
        except Exception as exc:  # noqa: BLE001 - se informa, no se oculta
            fallos.append((nombre, str(exc)))
    return ok, fallos
