---
name: daily-client-presentation-prep
description: Revisa reuniones de Sofia Garcia (el día en curso + los 2 siguientes días hábiles), genera la presentación GB Advisors de 9 slides del cliente desde vTiger con ajuste por agenda y guarda el HTML en la carpeta Reportes Claude - Presentation. Se dispara a mano desde el Panel de Automatizaciones (tarjeta "Rutinas automáticas (Sofia)").
---

Objetivo: en cada corrida, revisar si Sofia Garcia (sofia.garcia@gb-advisors.com) organizó reuniones dentro de la VENTANA DE TRES DÍAS (el día en curso más los 2 siguientes días hábiles) en el calendario de Outlook de el responsable comercial Quintero (<VTIGER_USERNAME del .env>), y dejar lista la presentación del cliente correspondiente en la carpeta de trabajo de el responsable comercial.

DEFINICIÓN DE LA VENTANA DE TRES DÍAS (importante, cambiada el 9-sep-2026 a pedido de el responsable comercial; antes era solo hoy + el próximo día hábil): la ventana es el DÍA EN CURSO más los DOS SIGUIENTES DÍAS HÁBILES, saltando sábados y domingos. Según el día en que se corra:
- LUNES → lunes, martes y miércoles.
- MARTES → martes, miércoles y jueves.
- MIÉRCOLES → miércoles, jueves y viernes.
- JUEVES → jueves, viernes y el LUNES siguiente.
- VIERNES → viernes, el LUNES y el MARTES siguientes.
- SÁBADO o DOMINGO (posible ahora que la rutina se dispara a mano): ese día no es hábil y no aporta reuniones; la ventana son los 2 siguientes días hábiles, es decir LUNES y MARTES.
Ejemplo: corriendo un viernes, se revisan las reuniones del viernes (hoy) y del lunes. Corriendo un martes, se revisan martes (hoy) y miércoles.

Para cada reunión encontrada dentro de la ventana de tres días que todavía no tenga una presentación reciente para ese cliente, crear la presentación comercial de GB Advisors del cliente correspondiente desde vTiger, con un ajuste ADICIONAL basado en la agenda de esa reunión, y exportarla como HTML listo para imprimir. EXCEPCIÓN IMPORTANTE: las reuniones sobre monday.com NO se procesan y NO se genera HTML para ellas, aunque las haya organizado Sofia Garcia — se saltan silenciosamente. REGLA DE NO DUPLICAR (ventana de 7 días, pedido explícito de el responsable comercial): si ya existe cualquier archivo de ese mismo cliente creado hace menos de 7 días, NO se regenera ni se duplica — se salta silenciosamente. Si no existe ninguna reunión de Sofia Garcia en toda la ventana de tres días (o todas las que hay son de monday.com o ya tienen presentación reciente), terminar sin hacer nada más (no generar nada, no notificar).

ENTREGA: el resultado de esta tarea es el archivo HTML guardado en la carpeta de el responsable comercial. NO se sube nada a SharePoint/OneDrive, NO se asigna la reunión a ningún vendedor y NO se envía ningún correo. Esta tarea es de solo lectura hacia Outlook y vTiger; lo único que escribe son los archivos HTML en disco.

Pasos:

1. Determinar las TRES fechas de la ventana (hora de Venezuela/VET UTC-4) según la definición de arriba: el día en curso y los 2 siguientes días hábiles. Luego buscar en el calendario de Outlook los eventos organizados por sofia.garcia@gb-advisors.com usando outlook_calendar_search (organizer="sofia.garcia@gb-advisors.com"), con un rango que cubra desde el inicio del PRIMER día de la ventana hasta el final del ÚLTIMO (VET). Nota: cuando la ventana cruza el fin de semana (jueves y viernes), ese rango abarca también sábado y domingo por conveniencia de la búsqueda, pero solo interesan las reuniones que caigan en los tres días de la ventana — descartar cualquier evento de sábado o domingo. Si no aparece ningún evento en ese rango, detener aquí — no continuar con los pasos siguientes.

2. Para cada evento encontrado dentro de la ventana de tres días:
   a. Leer su detalle completo (asunto, descripción/agenda, ubicación, asistentes) usando read_resource sobre el URI del evento. De ahí extraer:
      - El nombre de la organización/cliente (por asistentes externos, asunto o cuerpo del evento). El nombre que aparece en el asunto/cuerpo del correo de invitación tiene prioridad sobre el dominio de correo de los asistentes si hay discrepancia entre ambos (puede haber contactos en vTiger mal vinculados a otro dominio).
      - El producto/manufacturero de la reunión (por asunto, cuerpo/agenda o producto identificado).
      - Los temas/puntos de agenda mencionados, para usarlos en el ajuste adicional (paso 6 más abajo).
      - La fecha real de la reunión, que es la que se usa para nombrar el archivo.
   b. EXCLUSIÓN monday.com: si la reunión es sobre monday.com (detectado por el asunto, el cuerpo/agenda o el producto identificado en el paso 2a — p. ej. "monday.com", "monday.com Enterprise", "monday CRM/Work Management"), NO generar nada para ese evento: saltarlo silenciosamente y continuar con el siguiente, aunque la haya organizado Sofia Garcia. Dejar constancia en el mensaje final de que se saltó por ser de monday.com.
   c. CHEQUEO DE DUPLICADO POR CLIENTE CON VENTANA DE 7 DÍAS (regla explícita de el responsable comercial): ANTES de generar nada, revisar en la carpeta `<RAIZ_DEL_PROYECTO>/Reportes Claude - Presentation` (y también en la carpeta de outputs) si ya existe CUALQUIER archivo de ese mismo cliente — patrón `{nombre-cliente}-*.html`, SIN importar la fecha que lleve en el nombre — cuya fecha de creación/modificación sea de hace menos de 7 días (menos de 7×24 horas respecto al momento de la corrida). Si existe al menos uno reciente (< 7 días), este evento se considera ya cubierto: saltarlo silenciosamente y continuar con el siguiente, sin regenerar ni duplicar. Esto incluye el caso obvio de un archivo con el patrón exacto `{nombre-cliente}-{YYYY-MM-DD}.html` de esta misma reunión (que también se salta). Verificar la antigüedad por la fecha real del archivo en disco (mtime), no solo por la fecha del nombre. Solo si NO hay NINGÚN archivo de ese cliente creado en los últimos 7 días se procede a generar. Para comparar por cliente, normalizar el nombre a slug (minúsculas, guiones) igual que al nombrar el archivo, y considerar coincidencia si el slug del cliente aparece en el nombre del archivo. IMPORTANTE: la subcarpeta `_generador-deck` NO es un entregable, ignorarla en este chequeo.
   d. Si NO es de monday.com y NO existe ningún archivo reciente (< 7 días) de ese cliente, continuar con los pasos 3–8 para ese cliente/evento.

3. Autenticar contra vTiger con vtiger_authenticate usando:
   - username: el de la variable VTIGER_USERNAME del .env
   - accessKey: el de la variable VTIGER_ACCESS_KEY del .env

4. Buscar la cuenta en vTiger: `SELECT * FROM Accounts WHERE accountname LIKE '%[cliente]%' LIMIT 5` usando el nombre identificado en el paso 2a. Si hay varios resultados, elegir el que mejor coincida por nombre exacto con el cliente mencionado en el asunto/cuerpo de la reunión (esta es una corrida desatendida, no hay usuario disponible para desambiguar — elegir el mejor match por nombre y continuar, incluso si el dominio de correo de los asistentes no coincide con el de la cuenta).

5. GENERAR EL DECK CON EL GENERADOR FIJO (regla explícita de el responsable comercial, agosto 2026): el generador vive en `<RAIZ_DEL_PROYECTO>\Reportes Claude - Presentation\_generador-deck\` (build_deck.py + template.html + assets + README.md + ejemplo.json). Leer su README.md completo y seguirlo al pie de la letra: NO escribir ni editar el HTML a mano — se arma un archivo `datos.json` con el contrato exacto que exige el validador (ver README.md y `ejemplo.json` como referencia) y se corre `python build_deck.py datos.json salida.html` desde esa misma carpeta. Esto garantiza que todos los decks salgan con la misma maqueta. Si el generador devuelve un error de validación, corregir el JSON — nunca rodear el error escribiendo HTML manualmente. Si la carpeta `_generador-deck` no existe, dejar constancia clara del problema en el mensaje final y NO improvisar una maqueta distinta.

   La estructura es de **9 slides**: 1) portada, 2) su recorrido con GB Advisors, 3) su configuración actual, 4) GBS preliminar, 5) mapa de adopción, 6) novedades del producto, 7) casos de uso por industria, 8) oportunidades de expansión, 9) próximos pasos (vacía).

   5a. **Slide 3 — Su configuración actual** (permanente a partir de julio 2026, pedido explícito de el responsable comercial): consultar los Assets de la cuenta en vTiger (`SELECT * FROM Assets WHERE account = '[account_id]' LIMIT 20`, o por `serialnumber LIKE '%[dominio]%'`) y llenar las dos tarjetas `config_cards` del JSON con, como mínimo:
      - Producto/plan activo (`assetname`)
      - Agentes con licencia y precio por agente, leído de `cf_assets_licensedetails` (ej. "6 agents @40" → mostrar como "6 agentes @ $40 USD/agente"); si el mismo campo indica licencias incluidas sin costo (ej. "8 Field Service Management (free)"), mostrarlas como fila propia, porque suelen ser la oportunidad más valiosa de la sesión
      - Cambio reciente del contrato (ampliación o reducción) con su fecha, usando `description` / `cf_assets_operation` / `cf_assets_operationdate`
      - A través de quién se factura: campo `cf_assets_customerpaidto` (ej. "GB Advisors" vs. el fabricante directo)
      - Idioma de la plataforma (`cf_assets_languange`)
      - Modalidad de contrato (`cf_assets_paymentfrequency`)
      - Inicio de la relación (`datesold` o `cf_assets_licensereceiveddate`)
      - Período vigente desde (`dateinservice`) y próxima renovación (`cf_917` o el campo de fecha de renovación equivalente)
      - Si `assetstatus` aparece como "Out-of-service" o inconsistente con una renovación vigente, señalarlo en `config_note` como punto "a confirmar en vivo" con el cliente, y reflejarlo también como barrera en el GBS preliminar (slide 4).
      - **Excepción puntual a la regla de no mostrar montos** (ver Notas): esta lámina SÍ puede incluir el precio por agente tal como aparece en License Details, porque es información que el cliente ya conoce sobre su propio contrato. Esta excepción aplica únicamente a esta lámina — el resto del deck sigue sin mostrar MRR total, ARR, ni montos de historial de deals.

6. AJUSTE ADICIONAL (nunca reemplazar la estructura de 9 slides): incorporar los temas específicos de la agenda de esa reunión (del paso 2a) en el bloque `agenda` del JSON (la caja "alineación con la agenda" del slide 4) y como puntos suplementarios conectados a los slides de GBS/casos de uso/oportunidades (4, 7, 8). La estructura base de 9 slides debe quedar intacta; esto se agrega encima.

7. El generador ya produce un único archivo HTML autocontenido de 1280x720 por slide, con reglas `@media print` y `page-break-after` para imprimir con Ctrl+P sin reajustes, y hace un self-check antes de escribir (etiquetas balanceadas, 9 slides, cierre en `</html>`, pies de página, reglas de impresión). Si `build_deck.py` imprimió `OK`, el archivo está sano y NO hay que revalidarlo a mano ni "arreglarlo" editándolo.

8. Guardar el HTML con nombre `{nombre-cliente}-{YYYY-MM-DD}.html` (usando la fecha real de esa reunión, sea cual sea de los tres días de la ventana) directamente en la carpeta `<RAIZ_DEL_PROYECTO>/Reportes Claude - Presentation` — esa es la carpeta donde el responsable comercial quiere ver todos los archivos generados, sin necesidad de pedirlo cada vez. Este paso es la entrega final del evento: no subir el archivo a ningún otro destino ni notificar a nadie por correo.

9. Repetir los pasos 2–8 para cada evento restante de la ventana de tres días que no haya sido saltado por ser de monday.com o por tener ya una presentación reciente (< 7 días) del cliente.

10. Terminar la corrida. La notificación de finalización de la tarea programada es la señal para que el responsable comercial sepa que los archivos están listos en `<RAIZ_DEL_PROYECTO>/Reportes Claude - Presentation`.

Notas:
- La rutina se dispara a mano desde el Panel de Automatizaciones, así que puede correrse cualquier día, incluido fin de semana. La ventana siempre salta sábados y domingos: un viernes se prepara viernes + lunes + martes; un jueves, jueves + viernes + lunes; de lunes a miércoles, los tres días seguidos.
- MAQUETA FIJA (regla explícita de el responsable comercial, agosto 2026): todos los decks deben salir con la misma forma. Se usa siempre el generador fijo `_generador-deck/build_deck.py`. No usar el skill viejo `gb-presentation-full-capability` para esta tarea, y no escribir el HTML a mano bajo ninguna circunstancia.
- ENTREGA EN CARPETA (revertido en agosto 2026, pedido explícito de el responsable comercial): la presentación se entrega únicamente como archivo en `<RAIZ_DEL_PROYECTO>/Reportes Claude - Presentation`. Se eliminó la subida a SharePoint/OneDrive y se eliminó la asignación en rotación con envío de correos a vendedores. No usar sharepoint_upload_file, outlook_send_mail ni outlook_forward_mail en esta tarea. Los archivos `ausencias.txt` y `rotacion_estado.txt` ya no se leen ni se escriben.
- Las reuniones sobre monday.com se saltan siempre y no generan HTML, aunque las haya organizado Sofia Garcia (regla explícita de el responsable comercial). Solo se dejan registradas como "saltadas por ser de monday.com" en el mensaje final.
- REGLA DE NO DUPLICAR (ventana de 7 días): el chequeo del paso 2c es por CLIENTE, no por nombre-fecha exacto. Si ese cliente ya tiene cualquier presentación creada hace menos de 7 días (según la fecha real del archivo en disco), NO se genera otra — se salta silenciosamente. Solo se regenera si el archivo más reciente de ese cliente tiene 7 días o más.
- Si la autenticación a vTiger falla o no se encuentra la cuenta para alguno de los eventos, dejar constancia clara del problema en el mensaje final de la corrida para ese cliente puntual, y continuar con los demás eventos del rango sin bloquear todo el proceso.
- Si la agenda de un evento está vacía o no aporta nada útil, igual generar la presentación estándar sin el bloque `agenda` para ese cliente, y aclararlo en el mensaje final.
- No incluir MRR/ARR total de la cuenta, montos de historial de deals, IDs de vTiger, ni datos internos en el contenido orientado al cliente — **salvo la excepción explícita del paso 5a** (precio por agente en la lámina "Su configuración actual", que el cliente ya conoce). El generador bloquea automáticamente MRR, ARR e IDs de vTiger.
- Si por alguna razón la carpeta `<RAIZ_DEL_PROYECTO>/Reportes Claude - Presentation` no se puede escribir en una corrida dada, dejarlo indicado en el mensaje final — nunca bloquear la generación del archivo por esto.
- El mensaje final de cada corrida debe indicar claramente, y decir primero cuáles son los tres días de la ventana, por cada evento encontrado en ella: si se generó un HTML nuevo y dónde quedó guardado, o si se saltó (por presentación reciente < 7 días o por ser de monday.com), o si hubo algún problema.