---
description: Detecta menciones de vendedores sobre organizaciones de Sofía con señal de deal, publica el comentario en vTiger y reasigna el Assigned To (skill deteccion-reasignacion-orgs-sofia)
---

Sigue exactamente las instrucciones del archivo `<RAIZ_DEL_PROYECTO>/Claude/Scheduled/deteccion-reasignacion-orgs-sofia/SKILL.md` (ignora el frontmatter YAML, arranca desde la primera línea de instrucciones reales). Trátalo como si el responsable comercial te lo hubiera pedido explícitamente ahora mismo — este comando ES esa petición explícita.

Recordatorios clave del skill (actualizado 2026-08-27):
- La reasignación del "Assigned To" SÍ funciona por API vía `vtiger_revise` — hazla directo, no la dejes pendiente para que el responsable comercial la haga a mano.
- Las menciones en el comentario deben crearse YA con markup `<a class="mention">@Token</a>` (el texto plano "@Token" no se convierte solo), y con `is_private: "1"` y `publish_to: ["Users"]` explícitos para no exponer la nota al portal del cliente.
- La condición 3 (reunión pasada confirmada en calendario) sigue siendo obligatoria por defecto. Si el responsable comercial aprueba explícitamente saltarla para un caso puntual en el mismo pedido que disparó este comando, procesa ese caso igual sin la reunión.
