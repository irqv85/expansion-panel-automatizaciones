"""Asignar el producto a los farmings que no lo tienen.

Dos pasos, como el boton: LISTAR (solo lectura) y AÑADIR (escribe en vTiger).
Esta logica no depende de la pantalla, asi se puede probar sin tocar el CRM.

DETECCION AUTOMATICA, y hasta donde se puede confiar en ella. Medido el
9-oct-2026 contra 1.453 farmings de vTiger cuyo producto ya habia puesto una
persona (se esconde el producto, se detecta, se compara):

    el nombre trae el producto exacto   356 casos   96,3 % de aciertos   ALTA
    la descripcion menciona un producto 452 casos   73,5 %                baja
    los assets de la organizacion       470 casos   45,3 %                baja

Por eso hay dos niveles. Solo la confianza ALTA se escribe con el clic; la baja se
ofrece como sugerencia y la acepta la persona. Un producto equivocado en el CRM es
peor que uno vacio: el campo ya parece lleno y nadie vuelve a mirarlo.

Por que los assets rinden tan mal: un farming se hace para VENDER un producto, y
los assets son lo que el cliente YA tiene. Una cuenta con Freshdesk tiene farming
de Freshdesk Omni, o incluso de monday CRM. Y "CSM" es una categoria: las propias
personas eligieron Freshdesk Omni (67), Freshdesk Support Desk (58), monday CRM
(25) y Freshdesk (18) para farmings con ese mismo nombre.

Se probo tambien leer una frase de destino en la descripcion ("upsell from X to
Y"): acerto 40 % en 20 casos, asi que se descarto en lugar de afirmar «se vende Y»
con una certeza que no tiene.
"""
import html
import re

import vtiger_api as v

MODULO = "vtcmfarming"
CAMPO = "cf_vtcmfarming_product"

ALTA = "alta"
BAJA = "baja"

_CAMPOS = ("id, fld_vtcmfarmingname, cf_vtcmfarming_farmingstage, cf_vtcmfarming_product, "
           "cf_vtcmfarming_organization, cf_vtcmfarming_description, "
           "cf_vtcmfarming_memorydetails")


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


def _norm(texto):
    return re.sub(r"[^a-z0-9]+", " ", texto.lower()).strip()


def _contiene(texto, clave):
    return re.search(r"(?<![a-z0-9])" + re.escape(clave) + r"(?![a-z0-9])", texto) is not None


def sugerir(nombre, pares):
    """Producto que el NOMBRE del farming trae completo, o None.

    Exige el producto entero y como palabra suelta: asi "Halo" no se confunde con
    "Halogen". Si dos productos encajan, gana el mas largo ("ServiceNow - ITSM"
    antes que "Servicenow")."""
    texto = _norm(nombre)
    mejor = None
    for crudo, mostrar in pares:
        clave = _norm(mostrar)
        if len(clave) < 3:
            continue
        if _contiene(texto, clave) and (mejor is None or len(clave) > len(mejor[1])):
            mejor = (crudo, clave)
    return mejor[0] if mejor else None


def _mencionado(textos, claves):
    """Unico producto que aparece en el texto, o None si hay cero o varios.

    Se quitan los que son prefijo de otro mas largo tambien mencionado:
    "Freshdesk" dentro de "Freshdesk Omni 2026" no es una segunda mencion."""
    encontrados = {}
    for texto in textos:
        t = _norm(texto)
        if not t:
            continue
        for clave, crudo, mostrar in claves:
            if _contiene(t, clave):
                encontrados[crudo] = mostrar
    nombres = {c: _norm(m) for c, m in encontrados.items()}
    sobran = {a for a, na in nombres.items()
              if any(b != a and nb.startswith(na) for b, nb in nombres.items())}
    quedan = {c: m for c, m in encontrados.items() if c not in sobran}
    if len(quedan) == 1:
        return next(iter(quedan.items()))
    return None


def detectar(nombre, assets, pares, textos=()):
    """Producto de un farming: (crudo, evidencia, confianza).

    Si no hay base devuelve (None, motivo, None): NO adivina.

    Orden de evidencia, de mas a menos fiable (ver el encabezado del modulo):
      1. El nombre trae el producto exacto          -> ALTA
      2. La descripcion menciona un unico producto  -> baja
      3. Los assets de la organizacion dan uno solo -> baja
    """
    directo = sugerir(nombre, pares)
    if directo:
        mostrar = next(m for c, m in pares if c == directo)
        return directo, f"el nombre trae «{mostrar}»", ALTA

    claves = sorted(((_norm(m), c, m) for c, m in pares if len(_norm(m)) >= 4),
                    key=lambda t: -len(t[0]))

    unico = _mencionado(textos, claves)
    if unico:
        return unico[0], f"la descripcion menciona «{unico[1]}»", BAJA

    hallados = {}
    for asset in assets:
        texto = _norm(asset)
        for clave, crudo, mostrar in claves:
            if _contiene(texto, clave):
                hallados.setdefault(crudo, (mostrar, asset))
                break
    if len(hallados) == 1:
        crudo, (mostrar, asset) = next(iter(hallados.items()))
        return crudo, f"su asset «{asset}»", BAJA
    if not hallados:
        return None, "sin evidencia", None
    return None, "varios productos posibles: " + ", ".join(
        sorted(m for m, _ in hallados.values())), None


def listar_sin_producto(sesion, mi_id, pares):
    """Farmings asignados a `mi_id` sin producto, ya con la deteccion. SOLO LECTURA.

    Devuelve (filas, total). Cada fila lleva `sugerido` (producto crudo o None),
    `evidencia` (por que) y `confianza` (alta / baja / None)."""
    filas = v.consultar_todo(
        sesion,
        f"SELECT {_CAMPOS} FROM {MODULO} WHERE assigned_user_id = '{mi_id}'",
    )
    sin = [f for f in filas if not f.get(CAMPO)]

    # Assets de las organizaciones implicadas, en una sola consulta.
    orgs = sorted({f["cf_vtcmfarming_organization"] for f in sin
                   if f.get("cf_vtcmfarming_organization")})
    assets_por_org = {}
    if orgs:
        ids = ",".join(f"'{o}'" for o in orgs)
        for a in v.consultar_todo(
                sesion, f"SELECT assetname, account FROM Assets WHERE account IN ({ids})"):
            assets_por_org.setdefault(a["account"], []).append(a["assetname"])

    salida = []
    for f in sin:
        nombre = f.get("fld_vtcmfarmingname") or ""
        assets = assets_por_org.get(f.get("cf_vtcmfarming_organization"), [])
        textos = (f.get("cf_vtcmfarming_description") or "",
                  f.get("cf_vtcmfarming_memorydetails") or "")
        crudo, evidencia, confianza = detectar(nombre, assets, pares, textos)
        salida.append({
            "id": f["id"],
            "nombre": nombre,
            "etapa": f.get("cf_vtcmfarming_farmingstage") or "",
            "sugerido": crudo,
            "evidencia": evidencia,
            "confianza": confianza,
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
