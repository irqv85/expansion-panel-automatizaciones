"""El equipo de ventas, en un solo sitio y fuera del codigo.

Antes la lista de vendedores estaba escrita a mano en tres archivos distintos
(generar_reportes_next_step, Plan Compensacion y enviar_presentaciones_sofia),
cada uno con su propio formato. Para usar el panel con otro equipo habia que ir
a buscarlos uno por uno.

Ahora salen de `equipo.json`, en la raiz del proyecto. Ese archivo NO se versiona
(cada quien tiene el suyo): se parte de `equipo.example.json` y se edita.

    cp equipo.example.json equipo.json

Campos de cada vendedor:

    nombre       Como aparece en el campo "Assigned To" de vTiger. Tiene que
                 coincidir EXACTO: es la llave con la que se cruzan los reportes.
    correo       Para los envios por Outlook.
    alias        Otros nombres con los que el mismo vendedor aparece en vTiger,
                 por ejemplo si la cuenta se renombro. Opcional.
    hoja         Nombre de su hoja en el plan de compensacion, por si Excel no
                 admite el nombre completo (31 caracteres, nada de : \\ / ? * [ ]).
                 Opcional: por defecto se usa `nombre`.
"""
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
CONFIG = RAIZ / "equipo.json"
EJEMPLO = RAIZ / "equipo.example.json"


def _cargar():
    if not CONFIG.exists():
        sys.exit(
            f"ERROR: falta {CONFIG.name} en {RAIZ}.\n"
            f"       Copia {EJEMPLO.name} a {CONFIG.name} y pon ahi tu equipo.\n"
            f"       Los nombres tienen que coincidir exacto con el campo\n"
            f"       'Assigned To' de vTiger."
        )
    # Encoding explicito: Windows escribe en ANSI por defecto y un nombre con
    # acento rompe la lectura.
    with open(CONFIG, encoding="utf-8") as f:
        datos = json.load(f)

    vendedores = datos.get("vendedores") or []
    if not vendedores:
        sys.exit(f"ERROR: {CONFIG.name} no tiene ningun vendedor en 'vendedores'.")
    for v in vendedores:
        if not v.get("nombre"):
            sys.exit(f"ERROR: hay un vendedor sin 'nombre' en {CONFIG.name}.")
    return datos


_datos = _cargar()

#: Nombres tal como aparecen en vTiger. Es la lista que filtra todos los reportes.
TEAM_OWNERS = [v["nombre"] for v in _datos["vendedores"]]

#: alias en minuscula -> nombre canonico. Para cuando un mismo vendedor figura
#: en vTiger con mas de un nombre.
OWNER_ALIASES = {
    alias.lower(): v["nombre"]
    for v in _datos["vendedores"]
    for alias in v.get("alias", [])
}

#: nombre -> nombre de su hoja en el plan de compensacion.
REPS = {v["nombre"]: (v.get("hoja") or v["nombre"]) for v in _datos["vendedores"]}

#: nombre -> correo, para los envios por Outlook.
VENDEDORES = {
    v["nombre"]: v["correo"] for v in _datos["vendedores"] if v.get("correo")
}

#: Copia fija en TODOS los envios: los reportes de Calidad CRM y las
#: presentaciones de Sofia. Se llamaba cc_presentaciones, que enganaba porque
#: tambien lo usa Calidad CRM; se acepta el nombre viejo para no romper una
#: configuracion existente (7-oct-2026).
CC_EMAIL = (_datos.get("cc_fijo") or _datos.get("cc_presentaciones") or "").strip()

#: Texto que debe contener el nombre del export de vTiger para que este panel lo
#: considere. Sirve cuando hay mas de un equipo bajando reportes a la misma
#: carpeta de Descargas: sin esto, el panel toma el mas reciente de cada
#: categoria aunque sea del otro equipo, y los numeros salen de otra gente sin
#: que nada falle. Vacio = acepta cualquiera, que es el comportamiento de
#: siempre para quien tenga un solo equipo.
PATRON_REPORTES = (_datos.get("patron_reportes") or "").strip()

#: Como se llama este panel en su barra de titulo. Con mas de una instalacion en
#: la misma PC es lo que evita correr el reporte del equipo equivocado, que no
#: falla ni avisa: simplemente saca numeros de otra gente.
NOMBRE_PANEL = (_datos.get("nombre_panel") or "").strip()

#: Mostrar la tarjeta "Freshworks" (Forecast y Cadence Generator). Se apaga en
#: los equipos que no trabajan con Freshworks. Por defecto encendida, para no
#: cambiarle el panel a quien ya lo tenia.
MOSTRAR_FRESHWORKS = bool(_datos.get("mostrar_freshworks", True))

#: nombre -> id de usuario en vTiger (19x<numero>), solo de quienes lo tengan
#: configurado. Lo usa descargar_reportes_vtiger.py, que consulta la API por id
#: y no por nombre. Se saca abriendo una Organizacion del vendedor en vTiger y
#: mirando su assigned_user_id.
VTIGER_IDS = {
    v["nombre"]: v["vtiger_id"] for v in _datos["vendedores"] if v.get("vtiger_id")
}

#: Vendedores con su correo, en el orden en que se reparte el turno rotativo de
#: presentaciones. Lo consume Sofia - Schedule.py y la ventana de Vacaciones.
SELLERS = [
    {"name": v["nombre"], "email": v["correo"]}
    for v in _datos["vendedores"] if v.get("correo")
]
