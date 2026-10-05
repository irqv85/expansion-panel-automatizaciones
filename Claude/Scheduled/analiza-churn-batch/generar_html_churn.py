#!/usr/bin/env python3
"""
Generador FIJO del .html de un informe forense de churn, a partir de su .md.

Por que existe (14-sep-2026): el Paso 6 de la skill `analiza-churn` pedia que
el agente escribiera a mano un HTML de ~47 KB con la identidad de marca. Ese
paso fallo en la corrida de ese dia sobre Belize Electricity: quedo el .md sin
su .html y, como el Paso 4 de `analiza-churn-batch` exige que cada fila del
Excel enlace a un .html existente y al dia, la corrida entera se quedo sin
Excel. Ademas dejo basura en la carpeta de informes (12 trozos base64 de un
intento de reensamblar archivos).

Mismo criterio que build_deck.py con los decks: la maqueta no se escribe a
mano, se genera. El CSS sale de plantilla_churn.css, extraido del informe de
CIMA Group que ya estaba aprobado, asi que la salida se parece a lo que ya
habia y no inventa una paleta nueva.

Uso:
    python generar_html_churn.py <informe.md> [--salida <destino.html>]
    python generar_html_churn.py --todos      (genera los .html que falten)
"""
import html
import re
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "scripts"))
from paths import CHURN_INFORMES_DIR  # noqa: E402

PLANTILLA_CSS = Path(__file__).resolve().parent / "plantilla_churn.css"

# Logo de marca, el mismo SVG que traian los informes escritos a mano.
LOGO = (
    '<svg viewBox="0 0 110 86" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">'
    '<path fill-rule="evenodd" clip-rule="evenodd" d="M80 6H70V46V66H80V63.3243C82.9417 65.0261 '
    '86.3571 66 90 66C101.046 66 110 57.0457 110 46C110 34.9543 101.046 26 90 26C86.3571 26 '
    '82.9417 26.9739 80 28.6757V6ZM80 46C80 51.5228 84.4772 56 90 56C95.5228 56 100 51.5228 100 '
    '46C100 40.4772 95.5228 36 90 36C84.4772 36 80 40.4772 80 46Z" fill="var(--accent-2)"></path>'
    '</svg>'
)

# Campos de la Seccion 1 que van a las tarjetas de portada, en orden.
# Se toma el primero que exista, para tolerar informes con nombres algo
# distintos (los .md los escriben agentes y no siempre calcan la etiqueta).
TARJETAS = [
    (["MRR perdido"], "MRR perdido", None),
    (["ARR equivalente perdido", "ARR equivalente"], "ARR equivalente", None),
    (["Tipo de churn"], "Tipo de churn", None),
    (["Owner durante el churn", "Owner actual"], "Owner durante el churn", None),
    (["Tiempo con la cuenta"], "Tiempo con la cuenta", 160),
]

CLASE_VEREDICTO = {
    "no aplica": "good",
    "parcial": "warn",
    "aplica": "critical",
    "indetermin": "neutral",
}


def esc(t):
    return html.escape(t, quote=False)


def _inline(t):
    """Marcado en linea: negrita, cursiva, codigo y enlaces."""
    t = esc(t)
    t = re.sub(r"`([^`]+)`", r"<code>\1</code>", t)
    t = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<!\*)\*([^*\n]+)\*(?!\*)", r"<i>\1</i>", t)
    t = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2">\1</a>', t)
    return t


def _fila_tabla(linea):
    return [c.strip() for c in linea.strip().strip("|").split("|")]


def _es_separador(linea):
    return bool(re.fullmatch(r"\|?[\s:|-]+\|?", linea.strip())) and "-" in linea


def render_bloques(lineas):
    """Convierte un trozo de Markdown en HTML. Subconjunto deliberado:
    tablas, listas, citas, reglas y parrafos, que es todo lo que usan los
    informes de churn. No se trae una dependencia nueva por esto."""
    out = []
    i = 0
    n = len(lineas)
    while i < n:
        linea = lineas[i]
        bruto = linea.strip()

        if not bruto:
            i += 1
            continue

        # Tabla
        if bruto.startswith("|") and i + 1 < n and _es_separador(lineas[i + 1]):
            cab = _fila_tabla(bruto)
            i += 2
            cuerpo = []
            while i < n and lineas[i].strip().startswith("|"):
                cuerpo.append(_fila_tabla(lineas[i].strip()))
                i += 1
            out.append("<table><thead><tr>"
                       + "".join(f"<th>{_inline(c)}</th>" for c in cab)
                       + "</tr></thead><tbody>")
            for fila in cuerpo:
                fila += [""] * (len(cab) - len(fila))
                out.append("<tr>" + "".join(f"<td>{_inline(c)}</td>" for c in fila[:len(cab)]) + "</tr>")
            out.append("</tbody></table>")
            continue

        # Regla horizontal
        if re.fullmatch(r"-{3,}|\*{3,}|_{3,}", bruto):
            out.append('<div class="rule-sep"></div>')
            i += 1
            continue

        # Cita
        if bruto.startswith(">"):
            cita = []
            while i < n and lineas[i].strip().startswith(">"):
                cita.append(lineas[i].strip().lstrip(">").strip())
                i += 1
            out.append(f"<blockquote><p>{_inline(' '.join(cita))}</p></blockquote>")
            continue

        # Listas
        m_ul = re.match(r"[-*+]\s+(.*)", bruto)
        m_ol = re.match(r"\d+[.)]\s+(.*)", bruto)
        if m_ul or m_ol:
            etiqueta = "ul" if m_ul else "ol"
            patron = r"[-*+]\s+(.*)" if m_ul else r"\d+[.)]\s+(.*)"
            items = []
            while i < n:
                m = re.match(patron, lineas[i].strip())
                if not m:
                    if lineas[i].strip() and lineas[i].startswith(("  ", "\t")) and items:
                        items[-1] += " " + lineas[i].strip()
                        i += 1
                        continue
                    break
                items.append(m.group(1))
                i += 1
            out.append(f"<{etiqueta}>" + "".join(f"<li>{_inline(x)}</li>" for x in items) + f"</{etiqueta}>")
            continue

        # Subtitulo dentro de la seccion
        m_h = re.match(r"(#{3,6})\s+(.*)", bruto)
        if m_h:
            nivel = min(len(m_h.group(1)), 6)
            out.append(f"<h{nivel}>{_inline(m_h.group(2))}</h{nivel}>")
            i += 1
            continue

        # Parrafo: se juntan las lineas hasta un blanco o un bloque nuevo
        parrafo = []
        while i < n and lineas[i].strip() and not lineas[i].strip().startswith(("|", ">", "#")) \
                and not re.match(r"([-*+]|\d+[.)])\s+", lineas[i].strip()) \
                and not re.fullmatch(r"-{3,}|\*{3,}|_{3,}", lineas[i].strip()):
            parrafo.append(lineas[i].strip())
            i += 1
        if parrafo:
            out.append(f"<p>{_inline(' '.join(parrafo))}</p>")
        else:
            i += 1
    return "\n      ".join(out)


def partir_secciones(md):
    """Devuelve (titulo, [(numero, titulo, [lineas]), ...])."""
    lineas = md.splitlines()
    titulo = ""
    secciones = []
    actual = None
    for linea in lineas:
        m_h1 = re.match(r"#\s+(.*)", linea)
        m_h2 = re.match(r"##\s+(.*)", linea)
        if m_h1 and not m_h2:
            titulo = m_h1.group(1).strip()
            continue
        if m_h2:
            if actual:
                secciones.append(actual)
            cabecera = m_h2.group(1).strip()
            m_num = re.match(r"(\d+)[.)]\s*(.*)", cabecera)
            if m_num:
                actual = (m_num.group(1), m_num.group(2).strip(), [])
            else:
                actual = ("", cabecera, [])
            continue
        if actual:
            actual[2].append(linea)
    if actual:
        secciones.append(actual)
    return titulo, secciones


def tabla_campos(lineas):
    """Lee la tabla Campo|Valor de la Seccion 1 como diccionario."""
    campos = {}
    for linea in lineas:
        bruto = linea.strip()
        if not bruto.startswith("|") or _es_separador(bruto):
            continue
        celdas = _fila_tabla(bruto)
        if len(celdas) >= 2 and celdas[0].lower() not in ("campo", ""):
            campos[celdas[0]] = celdas[1]
    return campos


def buscar(campos, claves):
    for clave in claves:
        for k, v in campos.items():
            if k.lower().startswith(clave.lower()):
                return v
    return None


def clase_veredicto(texto):
    t = (texto or "").lower()
    for aguja, clase in CLASE_VEREDICTO.items():
        if aguja in t:
            return clase
    return "neutral"


def generar(ruta_md, ruta_salida=None):
    md = Path(ruta_md).read_text(encoding="utf-8")
    titulo, secciones = partir_secciones(md)

    campos = {}
    for num, _tit, lineas in secciones:
        if num == "1" or not campos:
            campos = tabla_campos(lineas)
            if campos:
                break

    organizacion = buscar(campos, ["Organizacion", "Organización"]) or titulo
    organizacion = re.sub(r"\s*\([^)]*\)\s*$", "", organizacion).strip()
    veredicto = buscar(campos, ["Veredicto"]) or ""
    confianza = buscar(campos, ["Confianza"]) or ""

    tarjetas = []
    for claves, etiqueta, corte in TARJETAS:
        valor = buscar(campos, claves)
        if not valor:
            continue
        if corte and len(valor) > corte:
            valor = valor[:corte].rsplit(" ", 1)[0] + "…"
        estilo = ' style="font-size:20px"' if len(valor) > 22 else ""
        tarjetas.append(
            f'<div class="stat"><p class="k">{esc(etiqueta)}</p>'
            f'<p class="v"{estilo}>{_inline(valor)}</p></div>'
        )

    cuerpo = []
    for num, tit, lineas in secciones:
        etiqueta_num = f'<span class="num">{esc(num.zfill(2))}</span>' if num else ""
        cuerpo.append(
            "  <section>\n"
            f'    <div class="section-head">{etiqueta_num}'
            f'<h2 class="h">{_inline(tit)}</h2><span class="rule"></span></div>\n'
            f'    <div class="prose">\n      {render_bloques(lineas)}\n    </div>\n'
            "  </section>"
        )

    css = PLANTILLA_CSS.read_text(encoding="utf-8")
    hoy = datetime.now().strftime("%d-%m-%Y")
    doc = f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Expediente {esc(organizacion)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,400;0,500;0,600;0,700;1,500&display=swap" rel="stylesheet">
<style>
{css}
.rule-sep{{height:1px;background:var(--line);margin:18px 0;border:0}}
</style>
</head>
<body>
<div class="wrap">

  <div class="cover">
    <div class="brandrow">{LOGO}</div>
    <h1>{esc(titulo or organizacion)}</h1>
    <p class="veredicto-portada"><span class="pill {clase_veredicto(veredicto)}">{_inline(veredicto) or "Sin veredicto"}</span>
       {f'<span class="pill small neutral">Confianza: {esc(confianza)}</span>' if confianza else ''}</p>
  </div>

  <div class="stats">
    {"".join(tarjetas)}
  </div>

{chr(10).join(cuerpo)}

  <footer>
    <p>Generado por <code>generar_html_churn.py</code> desde el informe <code>{esc(Path(ruta_md).name)}</code>
       &middot; GB Advisors &middot; documento confidencial de uso interno &middot; fuente: vTiger CRM &middot; {hoy}</p>
  </footer>

</div>
</body>
</html>
"""
    destino = Path(ruta_salida) if ruta_salida else Path(ruta_md).with_suffix(".html")
    destino.write_text(doc, encoding="utf-8")
    return destino, len(secciones)


def main():
    args = [a for a in sys.argv[1:]]
    if "--todos" in args:
        faltantes = [p for p in sorted(CHURN_INFORMES_DIR.glob("*.md"))
                     if not p.with_suffix(".html").exists()]
        if not faltantes:
            print("Todos los informes .md ya tienen su .html.")
            return
        for md in faltantes:
            destino, n = generar(md)
            print(f"Generado {destino.name} ({destino.stat().st_size} bytes, {n} secciones)")
        return

    if not args:
        print(__doc__)
        sys.exit(1)

    ruta_md = args[0]
    salida = None
    if "--salida" in args:
        salida = args[args.index("--salida") + 1]
    destino, n = generar(ruta_md, salida)
    print(f"Generado {destino} ({destino.stat().st_size} bytes, {n} secciones)")


if __name__ == "__main__":
    main()
