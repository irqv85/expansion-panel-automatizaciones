"""Cliente minimo de vTiger (webservice.php), con lectura y escritura separadas.

Existe aparte de descargar_reportes_vtiger.py a proposito: aquel script es un
generador de reportes de SOLO LECTURA y exige un mapa de usuarios para poder
importarse. Aqui hace falta escribir en el CRM, y eso se quiere en un modulo
chico, facil de auditar, donde la unica funcion que modifica algo se llama
`actualizar` y es imposible llamarla por accidente desde una consulta.

Las credenciales salen del .env (VTIGER_USERNAME, VTIGER_ACCESS_KEY, y
VTIGER_URL si la instancia no es la de siempre), igual que el resto del panel.
"""
import hashlib
import json
import os
import urllib.error
import urllib.parse
import urllib.request

import load_env  # noqa: F401  (carga .env al importarse)

BASE = (os.environ.get("VTIGER_URL") or "https://gbadvisors.od1.vtiger.com").rstrip("/")
if not BASE.endswith("webservice.php"):
    BASE += "/webservice.php"

TIMEOUT = 30


class VtigerError(RuntimeError):
    pass


# vTiger limita las llamadas por minuto y responde 429 cuando se pasa. Escribir N
# farmings son N llamadas seguidas, asi que sin esperar y reintentar el lote se
# cortaba a la mitad (9-oct-2026, medido al bajar datos de prueba).
ESPERAS = (4, 10, 25, 60)


def _abrir(pedir):
    import time
    for espera in ESPERAS + (None,):
        try:
            return json.loads(pedir().read())
        except urllib.error.HTTPError as exc:
            if exc.code != 429 or espera is None:
                raise
            time.sleep(espera)


def _get(params):
    url = BASE + "?" + urllib.parse.urlencode(params)
    return _abrir(lambda: urllib.request.urlopen(url, timeout=TIMEOUT))


def _post(params):
    data = urllib.parse.urlencode(params).encode()
    return _abrir(lambda: urllib.request.urlopen(BASE, data=data, timeout=TIMEOUT))


def login():
    usuario = os.environ.get("VTIGER_USERNAME", "")
    clave = os.environ.get("VTIGER_ACCESS_KEY", "")
    if not usuario or not clave:
        raise VtigerError("Faltan VTIGER_USERNAME o VTIGER_ACCESS_KEY en el .env.")
    ch = _get({"operation": "getchallenge", "username": usuario})
    if not ch.get("success"):
        raise VtigerError(f"getchallenge fallo: {ch}")
    token = ch["result"]["token"]
    acceso = hashlib.md5((token + clave).encode()).hexdigest()
    r = _post({"operation": "login", "username": usuario, "accessKey": acceso})
    if not r.get("success"):
        raise VtigerError(f"Login fallo: {r}")
    res = r["result"]
    return res["sessionName"], res.get("userId")


# ---------------------------------------------------------------- lectura

def consultar(sesion, vql):
    r = _get({"operation": "query", "sessionName": sesion, "query": vql})
    if not r.get("success"):
        raise VtigerError(f"Consulta fallo: {r.get('error')} -- {vql}")
    return r["result"]


def consultar_todo(sesion, select_from_where, lote=200, tope=5000):
    """Pagina sola: vTiger devuelve maximo 200 filas por consulta, y 'LIMIT' va
    con el offset primero."""
    salida, offset = [], 0
    while offset < tope:
        filas = consultar(sesion, f"{select_from_where} LIMIT {offset}, {lote};")
        salida.extend(filas)
        if len(filas) < lote:
            break
        offset += lote
    return salida


def describir(sesion, modulo):
    r = _get({"operation": "describe", "sessionName": sesion, "elementType": modulo})
    if not r.get("success"):
        raise VtigerError(f"describe fallo: {r.get('error')}")
    return r["result"]


def recuperar(sesion, id_registro):
    r = _get({"operation": "retrieve", "sessionName": sesion, "id": id_registro})
    if not r.get("success"):
        raise VtigerError(f"retrieve fallo: {r.get('error')}")
    return r["result"]


# -------------------------------------------------------------- escritura

def actualizar(sesion, registro_completo):
    """Guarda un registro. UNICA funcion de este modulo que modifica el CRM.

    Usa 'update', que reemplaza el registro entero, asi que recibe el registro
    completo tal como lo devolvio `recuperar` con el campo ya cambiado. Mandar
    solo el campo nuevo dejaria en blanco todos los demas."""
    r = _post({"operation": "update", "sessionName": sesion,
               "element": json.dumps(registro_completo)})
    if not r.get("success"):
        raise VtigerError(f"update fallo: {r.get('error')}")
    return r["result"]


def revisar(sesion, id_registro, campos):
    """Cambia SOLO los campos indicados, sin tocar el resto del registro.

    Es la variante parcial de `actualizar`: manda el id y los campos que cambian.
    Preferible para un cambio puntual, porque no hace falta leer el registro
    completo y no hay riesgo de dejar en blanco lo que no se mando."""
    elemento = {"id": id_registro}
    elemento.update(campos)
    r = _post({"operation": "revise", "sessionName": sesion,
               "element": json.dumps(elemento)})
    if not r.get("success"):
        raise VtigerError(f"revise fallo: {r.get('error')}")
    return r["result"]
