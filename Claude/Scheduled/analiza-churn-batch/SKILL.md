---
name: analiza-churn-batch
description: Corre el análisis forense de churn (skill analiza-churn) sobre TODAS las cuentas cuyo churn "aplica" según el dashboard (Churn Assets vs OA), en paralelo con varios agentes, y consolida el resultado en un Excel resumen (una fila por organización, con veredicto de responsabilidad del vendedor).
---

Eres un asistente para el responsable comercial Quintero (<VTIGER_USERNAME del .env>) en GB Advisors. Córrela SOLO cuando el responsable comercial lo pida explícitamente (ej. "analiza los churn", "corre el análisis de churn", "dame el resumen de churn de este mes").

=== QUÉ HACE (y por qué existe, agosto 2026) ===
El dashboard (`<RAIZ_DEL_PROYECTO>/Dashboard SMB-MM/Dashboard/Dashboard.html`, vista "Churn") ya calcula, cruzando los Excel `ChurnAssets` y `OA` de vTiger, qué organizaciones tienen un churn que **aplica** (no les queda otro asset activo en OA compensando) vs **no aplica** (les queda otro asset activo). Eso responde "¿es un churn real?", no "¿es culpa del vendedor?". Para lo segundo existe la skill `analiza-churn` (`<RAIZ_DEL_PROYECTO>/Churn Accounts/.claude/skills/analiza-churn/SKILL.md`), que hace un análisis forense completo de 19 secciones por cuenta y emite un veredicto de responsabilidad. Esta skill une ambas cosas: toma las organizaciones con churn aplicable y corre `analiza-churn` sobre cada una, en paralelo, y arma un Excel de una sola fila por cuenta para que el responsable comercial no tenga que abrir 5 informes de 19 secciones para saber a quién le aplica responsabilidad.

=== PASO 1. Obtener la lista de organizaciones con churn aplicable ===

1. Localiza en `la carpeta de Descargas del usuario (%USERPROFILE%\Downloads)` los Excel más recientes que matcheen `Mid Market Metricas 2025 - ChurnAssets_*.xlsx` y `Mid Market Metricas 2025 - OA_*.xlsx` (por fecha de modificación, el más nuevo de cada patrón). Si no hay ninguno, o el más reciente tiene más de 24h, avísale a el responsable comercial que corra "Generar" en el panel de Calidad CRM (o pídele que baje los Excel de vTiger a Descargas) antes de continuar, y detente.

2. Con Python (`openpyxl`, no lo tengas que reinstalar — ya está disponible en este entorno) replica exactamente la lógica de `churnValidationRows` del dashboard (función definida en `Dashboard.html`, línea ~2313 al momento de escribir esto — si cambió, léela de nuevo del archivo real en vez de confiar en este número de línea):
   - De `OA`, arma un mapa de assets activos (`Assets Current MRR > 0`) por organización (clave: `Organizations Organization ID` si existe, si no `Organizations Organization Name` en minúsculas).
   - Para cada fila de `ChurnAssets`, busca otros assets activos de la misma organización en ese mapa (excluyendo el propio asset que hizo churn por nombre). `applies = (no hay otros assets activos)`.
   - Filtra solo filas donde `Assets Previous MRR > 0` OR `Assets Current MRR == 0` OR `Assets Operation` (en minúsculas) `== 'cancellation'` — igual que el dashboard.
   - Quédate solo con `applies == True` (esas son las que vale la pena analizar a fondo; las que no aplican no son churn real, no requieren análisis forense).

3. Para cada organización que aplica, resuelve su cuenta real de vTiger: `vtiger_authenticate` (username: el de la variable VTIGER_USERNAME del .env, accessKey: el de la variable VTIGER_ACCESS_KEY del .env) y `SELECT id, accountname, account_no, assigned_user_id FROM Accounts WHERE account_no='<Organizations Organization ID>'`. El campo `Organizations Organization ID` del Excel (ej. `ACC1717`) es el `account_no` de Accounts; el `id` devuelto (`3x<N>`) es el que necesita `analiza-churn`. Si no matchea por `account_no`, intenta por `accountname LIKE '%<Organizations Organization Name>%'` y toma el mejor match.

=== PASO 2. Nunca reanalizar el mismo churn (regla explícita de el responsable comercial, agosto 2026) ===

Un churn es un evento histórico que pasa UNA sola vez — no como las presentaciones o los reportes periódicos de este proyecto, que sí se regeneran. Por eso acá el chequeo de duplicado NO es por ventana de días: es indefinido.

Antes de lanzar el análisis forense de una cuenta, revisa `<RAIZ_DEL_PROYECTO>/Churn Accounts/informes/` por CUALQUIER archivo `<slug>_<account_id>_*.md` de esa MISMA cuenta (mismo `account_id`, formato `3x<N>`), sin importar qué tan viejo sea. Si existe, NO la vuelvas a analizar bajo ninguna circunstancia: reutiliza ese informe (léelo para extraer los campos de la Sección 1 y 17) y anótala como "reutilizado, ya analizado el <fecha del informe>" en el resumen final.

Al extraer el campo `Esfuerzo comercial` de un informe reutilizado para el Excel resumen, aplica la misma regla del Paso 3: si el texto original de la Sección 1 mezcla vendedor y Renewals (ej. "NULO (vendedor) / RAZONABLE (Renewals)"), quédate solo con la parte del vendedor — el informe `.md` en sí no se toca, pero el Excel resumen nunca debe mostrar la parte de Renewals.

Excepción real (no por antigüedad del informe, sino porque es un evento distinto): si la cuenta vuelve a aparecer en `ChurnAssets` con un `Assets Asset Name` DIFERENTE al de cualquier informe ya guardado para ese `account_id` (o sea, un asset nuevo que no se había analizado antes hizo churn), eso sí es un churn nuevo y amerita un informe nuevo — pero el asset ya cubierto por un informe existente, nunca se reanaliza.

=== PASO 3. Análisis forense en paralelo (usa el tool Workflow) ===

Para cada cuenta SIN informe reciente (Paso 2), en paralelo (agentes independientes, uno por cuenta — usa varios agentes a la vez, no los corras en serie):

1. Cada agente debe: leer completo el archivo `<RAIZ_DEL_PROYECTO>/Churn Accounts/.claude/skills/analiza-churn/SKILL.md` y seguirlo AL PIE DE LA LETRA para la cuenta `3x<N>` asignada (autenticándose el mismo con las credenciales de arriba), incluyendo guardar el `.md` y el `.html` en `informes/` tal como indica esa skill (Paso 5-6 de esa skill). No le resumas tú el marco normativo en el prompt del agente — dile que LEA el archivo completo, es largo y las reglas importan (regla de antigüedad, regla de presentabilidad, notas técnicas de la API, etc.).
2. Además del informe completo, cada agente debe devolver (usa `schema` en el `agent()` del Workflow) un resumen estructurado: organización, account_id, owner/vendedor (nombre real, nunca ID), tiempo con la cuenta, tipo de churn (Full/Partial), MRR previo, MRR perdido, ARR perdido, causa probable, esfuerzo comercial (FUERTE/RAZONABLE/DÉBIL/NULO/INDETERMINADO), veredicto (una de las 5 clasificaciones permitidas de `analiza-churn`), confianza (ALTA/MEDIA/BAJA), resumen ejecutivo de 1-2 líneas, y la ruta del `.md` guardado.

**Guarda cada resumen apenas la cuenta termina, no al final (1-oct-2026, lección cara).** En cuanto un agente devuelva su resumen estructurado, escríbelo en `Claude/Scheduled/analiza-churn-batch/resumenes_cuentas.json` (una lista de objetos con `org_id`, `organization`, `tiempo`, `tipo_churn`, `causa`, `confianza`, `resumen_ejecutivo`, `verificacion`). Ese día el análisis forense de las 5 cuentas salió bien y aun así no hubo Excel: el resumen de todas vivía solo dentro de un script de un solo uso que la corrida alcanzó a escribir pero no a ejecutar. Con el JSON en disco, una corrida que se muere se retoma corriendo el Paso 4 y no se repiten horas de análisis.

   **El campo `esfuerzo comercial` es EXCLUSIVAMENTE del vendedor/owner de cuenta** (regla explícita de el responsable comercial, 27-ago-2026): a él no le interesa el desempeño de Renewals en esta tabla, eso ya está cubierto por separado en la Sección 12 del informe completo de cada cuenta. Si el agente que analiza la cuenta te devuelve algo mixto tipo "NULO (vendedor) / RAZONABLE (Renewals)", quédate solo con la clasificación del vendedor para este campo del Excel resumen — descarta la parte de Renewals, no la concatenes ni la pongas entre paréntesis.

3. Después de que cada análisis termine, lanza (en paralelo también, uno por cuenta recién analizada) un agente de VERIFICACIÓN independiente: dale el veredicto y el resumen ejecutivo del paso anterior (no el informe completo, para que juzgue con ojos frescos) y pídele que aplique el "Test final de responsabilidad" (las 7 preguntas de `analiza-churn`, sección homónima) contra la evidencia citada, y que responda si el veredicto le parece sostenible o si detecta una contradicción/salto lógico. Si el verificador señala un problema real, dilo en el resumen final como advertencia — no lo ocultes, pero tampoco reescribas el informe sin que el responsable comercial lo pida.

Para las cuentas SÍ reutilizadas (Paso 2), corre igual un agente de verificación contra el informe ya existente, para confirmar que el veredicto sigue siendo defendible (no hace falta rehacer el análisis, solo la relectura crítica).

**Cuando un verificador marque `sustainable=false` (aprendido en producción, 26-ago-2026, caso Destileria Serralles):** no lo tomes como automáticamente correcto ni lo repitas tal cual en el resumen para el responsable comercial. El verificador solo ve el resumen condensado, no el informe completo — en el caso de referencia confundió dos fechas distintas que el campo `cause` condensado mencionaba juntas (una señal de riesgo temprana y la confirmación final semanas después), y leyó una contradicción que no existía en el informe real. Antes de escribir la advertencia en el Excel o decírsela a el responsable comercial como un problema real, TÚ (el orquestador, no el agente verificador) relee la sección relevante del informe completo (`.md`) y confirma si la objeción se sostiene con el texto real. Si el verificador tenía razón, díselo a el responsable comercial tal cual. Si el verificador se equivocó por trabajar con el resumen comprimido, dilo también así de claro — no ocultes que hubo una objeción, pero no la repitas como válida sin haberla confirmado contra la fuente.

=== PASO 4. Armar el Excel resumen ===

**No escribas un script nuevo: corre el que ya está versionado.** Desde la carpeta del skill, en este orden:

```
python paso3_extraer_informes.py   # lee los .md de informes/ y arma informes_extraidos.json
python paso5_armar_excel_final.py  # arma el Excel con las 15 columnas
```

`paso3` sale de `cuentas_para_analizar.json` (lo deja el Paso 1) y busca el informe de cada cuenta por el slug del nombre, así que toma cualquier cuenta, no una lista fija; las que no tengan informe previo las deja en `cuentas_pendientes.json`, que es justo lo que hay que mandar al análisis forense del Paso 3. `paso5` combina eso con `churn_aplicable.json` (MRR, sumando todos los assets de la cuenta) y con `resumenes_cuentas.json` (los campos de juicio que dejaste en el paso anterior), y ya aplica colores, formato de moneda e hipervínculos. Si el Excel del día está abierto en Excel guarda al lado con sufijo y avisa, en vez de perder la corrida.

Revisa el resultado antes de darlo por bueno. El Excel queda con una fila por organización y estas columnas (en este orden):

`Organización | Org ID | Vendedor | Tiempo con la cuenta | Tipo de churn | MRR previo | MRR perdido | ARR perdido | Causa probable | Esfuerzo comercial | Veredicto | Confianza | Verificación | Resumen ejecutivo | Informe`

- `Informe`: hipervínculo real a la ruta local del archivo `.html` (regla explícita de el responsable comercial, 27-ago-2026: **nunca al `.md`** — abierto suelto se ve como texto plano sin tablas, "puros pipes"; el `.html` con la identidad de marca de GB Advisors que arma el Paso 6 de `analiza-churn` es el entregable presentable). Con openpyxl: `cell.value = "Ver informe"` (o el nombre de archivo), `cell.hyperlink = "file:///ruta/absoluta/al/archivo.html"`, y dale estilo de hipervínculo (`Font(color="0563C1", underline="single")`) para que se note que es clicable y abra el archivo local directo desde Excel. Antes de enlazar, confirma que el `.html` exista y esté tan al día como el `.md` (mismo timestamp o más reciente) — si el `.md` se corrigió después de generar el `.html`, regenera el `.html` primero.
- `Verificación`: "OK" si el agente verificador confirmó el veredicto, o su nota de advertencia si señaló algo.
- `Veredicto`: coloréalo (fill de la celda; si da problemas usa al menos el color de la fuente) según la clasificación — el semáforo no puede quedar sin color como pasó la última vez (regla explícita de el responsable comercial, 27-ago-2026):
  - 🟢 NO APLICA CHURN AL VENDEDOR / 🟢 NO APLICA — CONTROL/TIEMPO INSUFICIENTE → verde (fill `C6EFCE`, fuente `006100`)
  - 🟡 EVIDENCIA INSUFICIENTE → amarillo (fill `FFEB9C`, fuente `9C6500`)
  - 🟠 RESPONSABILIDAD COMERCIAL PARCIAL → naranja (fill `FCE4D6`, fuente `E26B0A`)
  - 🔴 APLICA RESPONSABILIDAD DE CHURN → rojo (fill `FFC7CE`, fuente `9C0006`)
- Aplica formato simple y legible: encabezados en negrita, columnas de MRR/ARR con formato de moneda, ancho de columna razonable (no hace falta réplica del branding de GB, esto es una herramienta de trabajo interna).
- Guarda el archivo en `<RAIZ_DEL_PROYECTO>/Churn Accounts/informes/resumen-churn-<YYYY-MM-DD>.xlsx` (fecha de hoy).

=== SECCIÓN DESPLEGAR (SÍ escribe: SharePoint) ===

Corre SOLO cuando el responsable comercial lo pida explícitamente (ej. "desplega el churn", "sube el churn a SharePoint"). Toma el Excel resumen más reciente, sube SOLO los .html a SharePoint EN PARALELO, y actualiza los links en el Excel (OPTIMIZADO 1-sep-2026: usa Workflow para paralelizar, más rápido).

PASO D1 -- Preparar la lista (script versionado, no lo hagas a mano):

```
python paso6_desplegar.py --preparar
```

Busca el `resumen-churn-*.xlsx` más reciente (entiende el sufijo `-N` que deja el
Paso 4 cuando el archivo estaba abierto en Excel), localiza la columna "Informe"
por su encabezado y no por posición fija, y deja en `despliegue_items.json` las
filas cuyo enlace sigue siendo local y cuyo `.html` existe. Las que ya tienen
enlace de SharePoint las omite, así que repetir el despliegue no resube lo que ya
está arriba.

Después ubica la carpeta de SharePoint: `sharepoint_folder_search(name="Churn
Accounts")` en el sitio "sales". Si no existe, créala. Guarda `driveId` e `id`.

PASO D2 -- Subir los .html EN PARALELO (un agente por archivo, usa Workflow):

Cada agente sube UN archivo y nada más:

1. Lee el `.html` completo con la tool `Read` y quítale el prefijo de número de
   línea que `Read` antepone a cada renglón.
2. `sharepoint_upload_file` con `driveId`, `parentItemId`, `filename`,
   `conflictBehavior="replace"`, y:
   - **`content` con el HTML tal cual, NO `contentBase64`.** El HTML es texto y
     la propia tool pide `content` para texto. Pasarlo en base64 lo engorda un
     33 por ciento y fue lo que tumbó el despliegue del 2-oct-2026 con el informe
     más grande. Nada de `base64 -w0`, ni archivos `.b64`, ni trocear.
   - `expectedBytes` con el tamaño exacto en bytes del archivo en disco
     (`wc -c < "<ruta>"`). El servidor rechaza la subida si lo que llega no
     decodifica a ese tamaño, así que una respuesta truncada falla en vez de
     publicar un informe recortado.
3. Si falla, reintenta UNA vez. Si vuelve a fallar, devuelve
   `{success: false, org, error: "<motivo literal>"}` y no insistas.
4. Si sale bien, devuelve `{success: true, org, webUrl}`.

**Si un archivo no te cabe en una sola respuesta, dilo en vez de mandar contenido
recortado.** Una subida truncada es peor que una que falta: queda publicada y
parece buena.

Sobre `expectedBytes`: un archivo con finales de línea CRLF llega a SharePoint
normalizado a LF y pesa un byte menos por renglón. Eso es normal y no es
corrupción (Belize: 62.018 bytes locales, 61.515 arriba, exactamente sus 503
CRLF). Si la diferencia no cuadra con el número de renglones, entonces sí hay
pérdida y hay que resubir.

PASO D3 -- Actualizar el Excel (script versionado):

Junta los resultados de los agentes en `despliegue_resultado.json`, una lista de
`{"org": "...", "success": true, "webUrl": "..."}`, y corre:

```
python paso6_desplegar.py --aplicar
```

Escribe los enlaces de SharePoint, guarda el Excel una sola vez, y a las que
fallaron les deja el enlace local a propósito para que se vea cuál quedó
pendiente. Si el Excel está abierto, guarda al lado con sufijo y avisa.

Es idempotente: si la corrida se muere después de subir, basta con volver a
correr `--aplicar` con el JSON de resultados. No hay que resubir nada.

PASO D4 -- Resumen final:
  Tabla Markdown: | Organización | Resultado | donde Resultado es "Desplegado: <webUrl>" / "Error al subir: <motivo>". Sé conciso.

=== PASO 5. Resumen final a el responsable comercial ===

Muestra en el chat la misma tabla (en Markdown) que quedó en el Excel, di cuántas cuentas se analizaron de nuevo vs. se reutilizaron, cualquier advertencia de los verificadores, y la ruta del Excel guardado.

Notas:
- Esta skill es de solo lectura hacia vTiger (los agentes de `analiza-churn` no modifican nada, ver regla de integridad #11 de esa skill).
- Si `analiza-churn` cambia de formato de reporte (secciones, nombre de archivo) en el futuro, esta skill debe seguir leyendo ese archivo tal cual esté, no una copia fijada aquí — por eso el Paso 3 le pide a cada agente leer el SKILL.md original en vez de embeber sus reglas aquí.
- No uses el tool Workflow para el Paso 4 (armar el Excel) — ese paso es determinístico y más simple/confiable hecho directo con Bash/Python, no como agente.
