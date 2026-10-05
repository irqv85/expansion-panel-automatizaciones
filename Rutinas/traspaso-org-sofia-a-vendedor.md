# Rutina: traspaso de organización de Sofia al vendedor

Detecta menciones de Vtiger hechas por el equipo Mid/SMB de el responsable comercial sobre organizaciones
propiedad de Sofia Garcia, y publica en la organización, en nombre de el responsable comercial, el comentario
que anuncia el traspaso de propiedad al vendedor y le pide a Sofia cerrar el farming.

## IDs verificados en Vtiger (2026-08-20)

| Persona | user_id | Handle de mención | Employee |
|---|---|---|---|
| Sofia Garcia | `19x474` | `@SofiaGarcia` | `52x14839785` |
| Vendedor 4 | `19x260` | `@NombreApellido` | `52x5169910` |
| Vendedora 1 | `19x431` | `@NombreApellido` | `52x12597063` |
| Vendedor 2 | `19x222` | `@NombreApellido` | `52x4142475` |
| Vendedora 3 | `19x425` | `@NombreApellido` | `52x12331938` |
| el responsable comercial Quintero (sesión) | `19x213` | `@IranQuintero` | |

El handle es nombre + apellido sin espacios, conservando apóstrofos y acentos.
`@NombreApellido` lleva apóstrofo, confirmado en comentario real de Rosangela Romero.

## Trigger

Correo de `message@crmalerts.vtiger.com` con asunto:

```
<Nombre del rep> mentioned you on <nombre del registro>
```

Solo cuentan los cuatro reps de la tabla. Se excluye `broadcasted message on`,
que no es una mención dirigida.

## Resolución registro a organización

NO buscar por nombre. El cuerpo HTML del correo trae el enlace directo al registro:

```
href="https://gbadvisors.od1.vtiger.com/view/detail?module=Accounts&id=13636321
      &activity=ModComments&selectedActivityId=15301341"
```

De ahí se saca `module`, `id` del registro y `selectedActivityId`, que es el ID del
comentario exacto. Esto es lo único confiable: hay organizaciones con nombres como
"fri" y deals con nombres duplicados o truncados que el matching por nombre no resuelve.

Prefijos de ID: `3x` Accounts, `5x` Potentials, `61x` vtcmfarming, `28x` ModComments.

Según el módulo, la organización sale de:

- `Accounts` → es la organización
- `Potentials` → `related_to`
- `vtcmfarming` → `cf_vtcmfarming_organization`

## Condiciones para publicar

Todas obligatorias:

1. La mención la hizo uno de los cuatro reps.
2. El CONTENIDO del comentario indica que el rep tuvo contacto real con el cliente y
   que hay deal: reunión sostenida, demo solicitada, "deal will open", oportunidad
   creada. Esta es la señal de traspaso.
3. `Accounts.assigned_user_id == 19x474` (Sofia es dueña de la organización).
4. No existe ya un comentario de traspaso en esa organización.

NO exigir que el rep tenga un deal abierto en la organización. En el caso "fri"
(2026-08-20) el deal todavía no existía, la vendedora escribió "Deal will open", y esa
condición habría bloqueado un traspaso legítimo.

Lo que se descarta por contenido: reportes de no-show ("client did not show up",
"customer does not appear at the session"), que son estatus, no traspasos.
También se descarta `broadcasted message on`, que no es mención dirigida.

## Acción: dos pasos

### 1. Comentario

Módulo `ModComments`, en inglés, sobre la organización:

- `related_to`: ID de la organización (`3xNNN`)
- `assigned_user_id`: `19x213` (el responsable comercial)
- `is_private`: `1`
- `publish_to`: `["Users"]`
- `commentcontent`:

```html
<div><p><a class="mention">@SofiaGarcia</a> <a class="mention">@{HandleVendedor}</a>:
{Nombre} will be taking over all activities on this account going forward, so ownership
of this organization will be transferred to {Nombre}. Sofia, please close the farming and
cancel any pending activities on your end. Thank you both.</p></div>
```

IMPORTANTE: `is_private` y `publish_to` hay que pasarlos explícitamente. Creando por
API, Vtiger los deja en `0` y `["Contacts"]`, que expone la nota al portal del cliente.
Los comentarios internos hechos desde la UI van con `1` y `["Users"]`.

### 2. Reasignar la organización

`vtiger_revise` sobre `Accounts` con `assigned_user_id` = user_id del vendedor.
Es reversible. El farming y las actividades pendientes los cierra Sofia, la rutina
no los toca.

## Restricciones operativas

- `ModComments` tiene `deleteable: false`. Un comentario publicado NO se puede borrar
  por API. Todo error es permanente y visible para el equipo. Verificar las cuatro
  condiciones antes de escribir.
- El módulo `Users` no es consultable por esta API (`ACCESS_DENIED`). Los user_id se
  resuelven vía `Employees.user_id`.
- `vtiger_retrieve_related` y `vtiger_add_related` requieren auth por header, no
  funcionan con `session_token`. Para listar comentarios usar
  `SELECT ... FROM ModComments WHERE related_to = '<id>'`.
- Sin verificar: si un `ModComments` creado por API dispara la notificación de mención
  de Vtiger. Revisar con Sofia y la vendedora si les llegó la del caso "fri".

## Validación 2026-08-20

13 menciones reales de los cuatro reps entre 08-18 y 08-20. Tres cayeron en
organizaciones de Sofia. Solo una era traspaso.

### Ejecutada: fri Guatemala `3x13636321`

Vendedora 1, 08/20: "Customer is requesting demo of FD Omni... The customer is
interesting in Bot and AI. Deal will open." Sofia llevaba desde el 03/07 con cinco
follow-ups a Carlos Hernandez sin una sola respuesta, farming en "Attempted To Contact".
la vendedora consiguió la conversación. Traspaso legítimo.

Hecho: comentario `28x15302073`, organización reasignada a `19x431`.

### Descartadas por contenido

| Organización | Rep | Contenido | Por qué no |
|---|---|---|---|
| INGEURBE SAS `3x1865430` | Vendedor 4 | "this client did not show up for the scheduled meeting" | No-show. Sofia tiene ciclo de demo vivo con Valeria Borda, slots 08/24, 08/26, 08/27. Deals solo de 2019 y 2020, cerrados. |
| T Maquinaria `3x12206679` | Vendedora 1 | "Customer does not appear at the session" | No-show. Los 6 deals cerrados, los dos de la vendedora perdidos en abril 2026. Sofia farmeando con Pedro Aguayo desde junio. |

Estas dos son la razón por la que el filtro no puede ser "cualquier mención sobre una
organización de Sofia". Con ese criterio se le habría pedido a Sofia retirarse de dos
cuentas que está trabajando, y el comentario no se puede borrar.

## Tarea programada

`traspaso-org-sofia`, cron `0 8 * * *` (diaria 8am local, el sistema le suma ~7 min de
jitter). Definicion en `%USERPROFILE%\.claude\scheduled-tasks\traspaso-org-sofia\SKILL.md`.

Ventana de busqueda: ultimas 24 horas de correo. Al correr diario, una mencion que llega
justo despues del disparo se procesa al dia siguiente.

Credenciales de Vtiger en `Rutinas\.vtiger-credentials` (texto plano en el disco de el responsable comercial).
Al rotar el access key en Vtiger, actualizar ese archivo.

Las tareas programadas solo corren con la app abierta. Si estaba cerrada a las 8am, corre
al siguiente arranque.
