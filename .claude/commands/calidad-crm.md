---
description: Regenera los reportes de Calidad CRM (Next Step) desde vTiger — limpia 'Emails Vendedores' y corre generar_reportes_next_step.py
---

Ejecuta el flujo equivalente al botón "Generar" de la tarjeta Calidad CRM del panel de automatizaciones (`Panel de Automatizaciones.py`, método `generar_reportes_ns`):

1. Borra todos los archivos existentes en `<RAIZ_DEL_PROYECTO>/Emails Vendedores` (solo archivos sueltos, no debería haber subcarpetas ahí). Esto evita que quede mezclado un reporte de una corrida anterior con la nueva.
2. Corre desde `<RAIZ_DEL_PROYECTO>/scripts`:
   `PYTHONIOENCODING=utf-8 python "generar_reportes_next_step.py"`
3. Reporta a el responsable comercial el resumen que imprime el script: deals y farmings revisados, cuántos con alerta de Next Step / producto, y para qué vendedores se generó archivo (y quién quedó sin alertas).

No envíes nada por correo en este comando: este comando solo genera. El envío se hace desde el panel, en la tarjeta Calidad CRM, con "Calcular envío" y luego "Enviar" — eso corre `enviar_reportes_next_step.py`, que manda por Outlook de escritorio con el HTML adjunto. No subas nada a SharePoint: el skill `enviar-reportes-next-step`, que subía los reportes y mandaba solo un enlace, se retiró el 9 de septiembre de 2026 porque el responsable comercial quiere el archivo adjunto.
