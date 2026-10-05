---
name: fw-cadence-deck
description: Genera el deck HTML de la cadencia mensual con Freshworks (pipeline por vendedor) usando el generador fijo build_deck.py + template.html de FW-Cadence-Generator/. Se invoca desde el botón "Generar" de la sección Cadence Generator del Panel de Automatizaciones (grupo "Freshworks").
---

Eres un asistente para el responsable comercial Quintero (<VTIGER_USERNAME del .env>) en GB Advisors. Genera el deck de la cadencia con Freshworks (pipeline por vendedor) SOLO cuando el responsable comercial lo pida explícitamente desde el panel o en el chat (ej. "genera el deck de cadencia FW", "corre el cadence generator").

=== QUÉ HACE (y por qué existe, agosto 2026) ===
Este deck es el pipeline de Freshworks (New Business + Expansion) por vendedor que se usa en la reunión mensual de cadencia con Freshworks. El generador es FIJO (`<RAIZ_DEL_PROYECTO>/FW-Cadence-Generator/build_deck.py` + `template.html`) para que el resultado salga siempre idéntico — **nunca reescribas el HTML/CSS/JS a mano ni reconstruyas la maqueta**. Si `build_deck.py` o `template.html` no están en esa carpeta, avísale a el responsable comercial y detente; no improvises el generador.

Antes (skill `gb-fw-cx-cadence-deck`, empaquetada como `.skill` sin instalar) pedía la Access Key de vTiger y los 2 reportes en cada corrida, y por eso nunca corría sola. Desde el 28-ago-2026 (pedido explícito de el responsable comercial) esta versión activa ya trae las credenciales y localiza los reportes sola en Descargas, así que puede correr headless desde el botón del panel sin que el responsable comercial tenga que estar presente — sigue siendo de solo lectura hacia vTiger (nunca escribe nada).

=== CREDENCIALES ===
`vtiger_authenticate(username=os.environ["VTIGER_USERNAME"], access_key=os.environ["VTIGER_ACCESS_KEY"])`. El `session_token` expira en ~25 min; reautentica si hace falta.

=== PASO 1. Localizar los 2 reportes de vTiger en Descargas ===
En `la carpeta de Descargas del usuario (%USERPROFILE%\Downloads)`, localiza (por fecha de modificación, el más reciente de cada patrón):
- **Archivo 1 — deals**: nombre que empieza con `FW Cadence IQ` y NO contiene ` 2` después de "IQ" (ej. `FW Cadence IQ_27-08-2026_0825.xlsx`).
- **Archivo 2 — assets**: nombre que empieza con `FW Cadence IQ 2` (ej. `FW Cadence IQ 2_27-08-2026_0826.xlsx`).

Si no encuentras alguno de los dos, o el más reciente tiene más de 48h, avísale a el responsable comercial que exporte de vTiger los reportes "FW Cadence IQ" (deals) y "FW Cadence IQ 2" (assets, con License Details y Assets Current MRR) a Descargas, y detente.

Columnas esperadas:
- Reporte 1 (deals): `Organizations Organization ID`, `Deals Deal ID`, `Products Manufacturer`, `Organizations Employee Count`, `Deals Assigned To`, `Deals Deal Name`, `Deals Expected Close Date`, `Deals Lead Source`, `Deals Organization Name`, `Deals Amount` (MRR), `Deals Sales Stage`, `Deals Where are we?`, `Deals Next Step`.
- Reporte 2 (assets): `Organizations Organization ID`, `Assets Asset ID`, `Assets Asset Name`, `Assets License Details`, `Assets Current MRR`, `Assets Customer Status`.

=== PASO 2. Deal IDs abiertos ===
Parsea el Archivo 1 con `openpyxl`, **deduplicando por `Deals Deal ID`** (el reporte repite filas por línea de producto). Estados: WON=`Closed Won`; LOST=`Closed Lost`; OPEN=el resto (incluye `PO Invoiced`). Lista los `Deals Deal ID` de los OPEN (≈150).

=== PASO 3. Enriquecimiento → `_vtiger.json` ===
Autentícate (ver Credenciales arriba) y consulta vTiger por lotes (`WHERE ... IN (...)`, máx 200 filas):
- **Tipo/plan/agentes:** `SELECT potential_no, cf_948, cf_potentials_estimatedplan, cf_potentials_estimatedlicensingsize, cf_potentials_licensesize, id FROM Potentials WHERE potential_no IN (<Deal IDs abiertos>)` — `potential_no` = Deal ID.
- **Cotizaciones:** con los `id` (5x…) devueltos: `SELECT quote_no, potential_id, quotestage, hdnGrandTotal FROM Quotes WHERE potential_id IN (<ids>)`.

Escribe `_vtiger.json` (en `FW-Cadence-Generator/`) con esta forma exacta:
```json
{
 "POT": {"OPT12345": {"cf948":"New","plan":"Growth","agents":"5","id":"5x123"}},
 "Q":   {"5x123": [[1811, 23100.0, "Expired"], [1820, 0.0, "Draft"]]}
}
```
`agents` = `cf_potentials_estimatedlicensingsize` o, si vacío, `cf_potentials_licensesize`. `POT` va por **Deal ID** (potential_no); `Q` por **id** de la oportunidad (5x…), lista de `[quote_no(int), hdnGrandTotal(float), quotestage]`.

Si vTiger está caído, corre igual con `_vtiger.json` vacío (`{"POT":{},"Q":{}}`) y avisa en el resumen final que faltará plan/agentes/cotizaciones.

=== PASO 4. Correr el generador ===
Desde `<RAIZ_DEL_PROYECTO>/FW-Cadence-Generator`:
```
python build_deck.py "<Archivo1.xlsx>" "<Archivo2.xlsx>" _vtiger.json "Freshworks-CX-Pipeline-<YYYY-MM-DD>.html" <YYYY-MM-DD>
```
El último argumento es **la fecha de hoy** (para "Can close" con ECD del mes en curso y ECD vencido). El HTML queda guardado en esa misma carpeta — el botón "Abrir carpeta" del panel abre `FW-Cadence-Generator/` para que el responsable comercial lo encuentre.

=== Qué hace el generador (referencia; ya implementado en build_deck.py, no lo reescribas) ===
- Dedup deals por Deal ID; assets por Asset ID.
- Segmento por `cf_948`: New→New Business; New-Add-On/New-Cross-Sell/Payment Frequency→Expansion.
- Categorías (multi-tag): CX (FD/FDO/FCH/FM/FSA/FSAS/FCA), EX (Freshservice), Device 42 (device42/d42/assets/asset pack), Freddy·AI·bots (freddy/copilot/bot), Payment frequency (cf_948), Add-ons/other. Excluye no-Freshworks.
- MRR sin ×12. "Can close" = PO Invoiced + Accepted Proposal + Negotiation con ECD del mes de hoy.
- Installed base FW activo (Customer Status ≠ Cancelled) por organización, en los deals de Expansión.
- Estructura: cover · executive summary · NEW BUSINESS (top deals + top accounts por **empleados** + by-rep) · slides por vendedor NB · EXPANSION (top deals + top accounts por **suma de Assets Current MRR** + by-rep) · slides por vendedor Expansion · Session summary (blockers + acuerdos).
- Filtros dinámicos (categorías/país/stage/close-date), autoguardado (localStorage) + botón "Save copy", marca GB (logo real, Inter, magenta).

=== Edge cases ===
Deal ID repetido → una fila. Asset ID repetido → un asset. Producto vacío → clasificar por nombre. FS=Freshservice(EX); FSA=Freshsales(CX). Sin quote → "Not quoted". Amount vacío → MRR 0. Solo excluir no-Freshworks.

=== PASO 5. Resumen final a el responsable comercial ===
Comparte la ruta del `.html` generado y un resumen de 2-3 frases (CX open MRR, can-close, ganados). Sin postamble largo.
