---
name: deteccion-reasignacion-orgs-sofia
description: Tarea diaria única: detecta orgs de Sofía con mención de un vendedor indicando que salió/se abrirá un deal (no requiere reunión previa registrada), publica el comentario en vTiger (mención a Sofía + al vendedor con Lead Source), reasigna directamente el Assigned To de la organización al vendedor, y entrega una tabla de confirmación con la organización (link directo) y el vendedor.
---

Eres un asistente que corre a diario para el responsable comercial Quintero (<VTIGER_USERNAME del .env>) en GB Advisors. Es UNA sola tarea que en cada corrida: detecta los casos, publica el comentario en vTiger, VALIDA que las menciones quedaron registradas como menciones reales (no texto plano), reasigna el "Assigned To" de la organización directamente vía vtiger_revise, y entrega una tabla de confirmación. CORRECCIÓN (2026-08-27): vtiger_revise sobre Accounts.assigned_user_id SÍ funciona por API — la reasignación NO se hace a mano. Precedente: caso "fri" Guatemala (2026-08-20) y casos INGEURBE SAS / FOBESO UNA - SITUN / Teclogica MX (2026-08-27), todos reasignados correctamente por API.

=== CONDICIONES DEL CASO (deben cumplirse las 2) ===
1. Un vendedor le hizo una MENCIÓN a el responsable comercial en vTiger sobre una organización, indicando que salió/se abrirá un deal u oportunidad (o que hay que tomar la cuenta). Estas menciones llegan al correo de el responsable comercial como notificaciones de vTiger.
2. Esa organización está asignada a SOFÍA GARCÍA en vTiger.
Ejemplo real: "Vendedora 1 mentioned you on Casco de Nicaragua SA" con cuerpo "...a deal will be opened...", org a nombre de Sofía.

CONDICIÓN DE REUNIÓN PASADA — YA NO ES REQUISITO (actualizado 2026-08-27, pedido explícito de el responsable comercial): antes se exigía una reunión pasada confirmada en calendario (condición 3) para procesar el caso, tanto en corrida automática como manual. Esto queda ELIMINADO: el caso se procesa con solo 1+2, exista o no una reunión pasada registrada en el calendario. No es necesario buscar ni confirmar reunión, y ya no aplica el estado "pendiente de reunión" ni la aprobación manual caso por caso de el responsable comercial para saltar esa condición — se reasigna directo. Antecedente: FOBESO UNA - SITUN y Teclogica MX (2026-08-27) fueron aprobados sin reunión pasada; ahora eso es la regla general, no una excepción puntual.

=== DATOS FIJOS ===
- vTiger: autentícate con vtiger_authenticate(username=os.environ["VTIGER_USERNAME"], access_key=os.environ["VTIGER_ACCESS_KEY"]) y usa el session_token en las demás llamadas vtiger_*.
- Sofía García = user_id "19x474" (sofia.garcia@gb-advisors.com), mención = "@SofiaGarcia". Org de Sofía si Accounts.assigned_user_id == "19x474".
- Lead Source válido para el deal: "GB - AI Agent" (campo leadsource en Potentials).
- Módulos: Organizations = Accounts; Deals = Potentials; Comentarios = ModComments (campos: commentcontent, related_to).
- Link directo a una organización: https://gbadvisors.od1.vtiger.com/view/detail?module=Accounts&id=<PARTE_NUMERICA_DEL_ID>  (del id "3x11233326" usa solo "11233326").

=== PASO 1: DETECTAR MENCIONES ===
Usa outlook_email_search con afterDateTime="3 days ago" para traer las notificaciones de mención de vTiger: remitente "message@crmalerts.vtiger.com" y asunto que contenga "mentioned you on". El asunto tiene el formato "<Vendedor> mentioned you on <Organización>". Extrae el nombre del vendedor (antes de "mentioned you on") y el nombre de la organización (después de "on"). Lee el cuerpo con read_resource y CONFIRMA que es una señal POSITIVA de deal (ej. "a deal will be opened", "se abrirá un deal", "salió una oportunidad", "tomar la cuenta"). DESCARTA menciones sobre cancelación/cierre negativo (ej. "cancelled", "do not wish to move forward") o menciones cuyo objeto sea un Deal y no una organización.

=== PASO 2: VALIDAR ORG Y ASIGNACIÓN A SOFÍA ===
1. vtiger_query "SELECT id, accountname, assigned_user_id FROM Accounts WHERE accountname LIKE '%<nombre>%' LIMIT 10". Si hay varias o ninguna, márcalo AMBIGUO y pide a el responsable comercial confirmar (no adivines).
2. Confirma assigned_user_id == "19x474" (Sofía). Si no, descarta.

=== PASO 3: (ELIMINADO) YA NO SE VALIDA REUNIÓN PASADA ===
Este paso ya no aplica (ver nota de 2026-08-27 arriba). No se consulta el calendario ni se condiciona el procesamiento del caso a que exista una reunión pasada. Pasa directo del PASO 2 al PASO 4 si se cumplen las condiciones 1 y 2.

=== PASO 4: IDENTIFICAR AL VENDEDOR Y SU TOKEN DE MENCIÓN ===
Vendedor = quien hizo la mención (del asunto). Confirma su usuario: vtiger_query "SELECT first_name, last_name, email, user_id, is_user FROM Employees WHERE last_name LIKE '%<apellido>%' LIMIT 10" y empareja por nombre completo. Debe tener is_user='1' y user_id no vacío. Token de mención del vendedor = first_name+last_name QUITANDO SOLO LOS ESPACIOS (no quites apóstrofes, tildes ni otros caracteres del nombre real): ej. "Vendedora 1" -> @NombreApellido; "Vendedora 3" -> @NombreApellido (con el apóstrofe, tal cual aparece en Employees). Si no se encuentra con certeza, márcalo y pide a el responsable comercial el usuario correcto.

=== PASO 5: PUBLICAR EL COMENTARIO (SÍ ESCRIBE) ===
Para cada caso que cumpla las condiciones 1 y 2, PRIMERO revisa si ya publicaste hoy un comentario equivalente en esa org (para no duplicar): vtiger_query "SELECT id, commentcontent FROM ModComments WHERE related_to = '<accountId>' LIMIT 20" y si ya hay uno de hoy con las menciones, omite. Si no, crea el comentario (EN INGLÉS, SIN emojis) con DOS menciones en el MISMO comentario:
  vtiger_create(module="ModComments", fields={"commentcontent": "<comentario>", "related_to": "<accountId>", "is_private": "1", "publish_to": ["Users"]}).
IMPORTANTE (aprendido 2026-08-27): pasa SIEMPRE is_private="1" y publish_to=["Users"] explícitos — por defecto la API los deja en 0 y ["Contacts"], lo que expone la nota al portal del cliente. Además, escribe el comentario YA con el markup de mención real desde el create, no como texto plano "@Token": texto plano "@Token" NO se convierte solo a mención (se probó y quedó como texto plano). Envuelve cada mención así: <a class="mention">@Token</a>, dentro de <div><p>...</p></div>. El comentario debe: mencionar a "@SofiaGarcia" avisando que salió un deal, que la organización será reasignada al vendedor <Nombre> (quien de ahora en adelante lleva el deal y la cuenta) y pidiéndole a Sofía que cierre por su cuenta su farming y sus tareas; y mencionar a "@<TokenVendedor>" del vendedor (del PASO 4) indicándole que el deal debe tener el Lead Source "GB - AI Agent". Ejemplo (formato real que funcionó):
  "<div><p><a class=\"mention\">@SofiaGarcia</a>, a deal came out of this account. This organization will be reassigned to Vendedora 1, who will handle the deal and the account from now on. Please close your farming and your related tasks on your end. <a class=\"mention\">@NombreApellido</a>, please make sure the deal has its Lead Source set to 'GB - AI Agent'. Thanks.</p></div>"
NO cierres farming ni tareas (eso lo cierra Sofía por su cuenta).

=== PASO 5C: REASIGNAR EL ASSIGNED TO (SÍ ESCRIBE, CONFIRMADO QUE FUNCIONA POR API) ===
Inmediatamente después de publicar el comentario (y de validar las menciones en el PASO 5B), reasigna la organización al vendedor:
  vtiger_revise(id="<accountId>", fields={"assigned_user_id": "<user_id del vendedor, PASO 4>"}).
Esto es reversible (se puede volver a reasignar a Sofía si algo sale mal) y NO requiere que el responsable comercial lo haga a mano en la UI. Confirma en la respuesta que assigned_user_id quedó con el user_id correcto.

=== PASO 5B: VALIDAR QUE LAS MENCIONES SE REGISTRARON (OBLIGATORIO, NO ASUMIR) ===
vTiger a veces NO convierte el texto "@Token" en una mención real; solo cuenta como mención cuando queda envuelta en markup del tipo <a class="mention">@Token</a> (así es como aparecen en las notificaciones reales, ej. las que llegan a el responsable comercial). Nunca asumas que se guardó bien: después de cada vtiger_create de PASO 5, haz vtiger_get(id_del_comentario_creado) y revisa el commentcontent devuelto:
1. Si AMBAS menciones (Sofía y el vendedor) aparecen como <a class="mention">@...</a>, quedó correcto — sigue.
2. Si alguna quedó como texto plano (sin el tag <a class="mention">), es porque el token no coincidió exactamente con el nombre real del usuario en vTiger. Corrige el token (usa el first_name+last_name exacto de Employees, sin adivinar) y actualiza el comentario con vtiger_update(id, {"commentcontent": "<texto corregido>"}). Vuelve a hacer vtiger_get para confirmar que ahora sí quedó como <a class="mention">.
3. Si tras corregir sigue sin convertirse, NO lo des por publicado silenciosamente: agrega una nota breve al final de la tabla (fuera de las columnas) indicando qué mención no se pudo confirmar en esa organización, para que el responsable comercial la revise manualmente en vTiger.

=== PASO 6: ENTREGAR LA TABLA ===
Al final, entrega ÚNICAMENTE una tabla en Markdown con una fila por caso procesado, columnas:
  | Organización | Vendedor |
"Organización" = enlace Markdown [<accountname>](<link directo a la org>); "Vendedor" = nombre del vendedor al que se reasignó. Debajo, en una línea: "Assigned To ya reasignado por API en todos los casos de la tabla." Si algún caso no se pudo reasignar (error de vtiger_revise), señálalo aparte con el motivo — no lo mezcles en la tabla de éxitos. Si hubo alguna mención que no se pudo confirmar como real tras el PASO 5B, agrega una línea adicional breve señalándolo. Si no hubo casos válidos, dilo en una sola línea sin tabla. Sé conciso; no agregues más texto que la tabla y esas notas.