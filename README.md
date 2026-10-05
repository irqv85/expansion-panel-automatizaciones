# Expansion - Panel de automatizaciones

Panel de escritorio que automatiza el trabajo comercial recurrente sobre vTiger CRM:
control de calidad del CRM, forecast, métricas de trimestre, planes de compensación,
análisis de churn y preparación de presentaciones de cliente.

Está pensado para un líder comercial que supervisa un equipo de vendedores. Lee los
reportes que exportas de vTiger, cruza los datos, y genera los Excel y correos que
antes había que armar a mano. Los envíos salen por tu Outlook de escritorio, así que
quedan en tus Elementos enviados como cualquier otro correo tuyo.

## Esta instalación

Es la de **Expansion Partners**, el equipo que trabaja con Humand, ServiceNow, Halo,
NinjaOne y Atera. Equipo configurado en `equipo.json`: María D'Amico, Valentina
Guedez y Miguel Rivero.

Comparte código con el panel de Freshworks, que es el mismo programa con otro
`equipo.json`. Lo único que cambia entre los dos es la configuración, así que una
mejora en uno se puede traer al otro con un `git pull` del remoto correspondiente.

**Tres cosas que importan por tener dos paneles en la misma PC:**

1. El título de la ventana dice **Expansion Partners**. Si dice otra cosa, estás en
   el panel equivocado. Correr un reporte con el equipo que no es no falla ni avisa:
   simplemente saca cifras de otra gente.

2. `patron_reportes` está en `"Expansion"`. El panel barre tu carpeta de Descargas y
   se queda con el export más reciente de cada categoría; sin ese filtro tomaría los
   de Freshworks, que conviven ahí. Ajústalo al prefijo real con el que nombres tus
   reportes en vTiger.

3. **María D'Amico está en los dos equipos este trimestre**, porque conserva deals de
   Freshworks de forma temporal. Sus cifras van a aparecer en los dos paneles. Al
   sumar totales de equipo o calcular compensación, no la cuentes dos veces.

## Sobre los fabricantes

El panel nació para un equipo de Freshworks y quedan rastros de eso, que no estorban
pero conviene conocer:

- Las columnas `Organizations Freshworks Tier` en los exports y `Freshworks ID` y
  `Freshwork Invoice` en el plan de compensación. Para Expansion vendrán vacías.
  Déjalas o quítalas de tus reportes de vTiger, da igual: los scripts no fallan si
  están vacías.
- La tarjeta **Cadence Generator** arma el deck de la cadencia con Freshworks. No
  aplica a Expansion mientras no exista una cadencia equivalente con Humand,
  ServiceNow, Halo, NinjaOne o Atera.

Todo lo demás (calidad de CRM, forecast, métricas, compensación, churn, GBS) es
agnóstico del fabricante: trabaja sobre deals, organizaciones, assets y farmings.

---

## Antes de empezar

Esto es lo que necesitas tener listo. Si falta algo, el panel te lo dice al arrancar.

| Requisito | Detalle |
|---|---|
| Windows | 10 u 11. El panel usa COM para hablar con Outlook, así que no corre en Mac ni Linux. |
| Python | 3.12 o 3.13, en el PATH. Comprueba con `python --version`. |
| Paquetes | `pip install openpyxl pypdf pillow pywin32 PySide6` |
| Outlook | De escritorio, instalado y con tu sesión iniciada. No sirve la versión web. |
| vTiger | Una cuenta con permiso de lectura sobre Deals, Organizations, Assets y Farmings. |
| Claude Code | Solo si vas a usar las rutinas desatendidas y las automatizaciones con IA. |

---

## Instalación

### 1. Clona el repositorio

```bash
git clone https://github.com/irqv85/expansion-panel-automatizaciones.git
cd expansion-panel-automatizaciones
```

Puedes ponerlo donde quieras: el proyecto deduce su propia ubicación y no lleva
rutas absolutas escritas. Si lo mueves de carpeta, de disco o de PC, sigue
funcionando sin editar nada.

### 2. Pon tus credenciales de vTiger

```bash
cp .env.example .env
```

Y edita `.env`:

```
VTIGER_USERNAME=tu.usuario@tuempresa.com
VTIGER_ACCESS_KEY=tu_access_key
VTIGER_URL=https://tuempresa.od2.vtiger.com
```

La **access key no es tu contraseña**. Se saca de vTiger, en *My Preferences*,
sección *Access Key*. Si alguna vez se te filtra, la rotas desde esa misma
pantalla y la anterior deja de servir al instante.

`.env` está en `.gitignore` y no se sube nunca. No pongas la clave en ningún otro
archivo: todos los scripts la leen de ahí.

### 3. Define tu equipo

```bash
cp equipo.example.json equipo.json
```

Y pon a tus vendedores:

```json
{
  "vendedores": [
    {
      "nombre": "Nombre Apellido",
      "correo": "nombre.apellido@tuempresa.com",
      "alias": [],
      "hoja": "Nombre Apellido"
    }
  ],
  "cc_presentaciones": "copia@tuempresa.com"
}
```

| Campo | Para qué sirve |
|---|---|
| `nombre` | **Tiene que coincidir exacto con el campo "Assigned To" de vTiger.** Es la llave con la que se cruzan todos los reportes. Un acento o un espacio de más y ese vendedor desaparece de los informes sin aviso. |
| `correo` | A dónde se le mandan sus reportes por Outlook. |
| `alias` | Otros nombres con los que el mismo vendedor aparece en vTiger, por ejemplo si se renombró la cuenta. Opcional. |
| `hoja` | Nombre de su pestaña en el plan de compensación. Úsalo si el nombre completo no cabe en Excel (31 caracteres, sin `: \ / ? * [ ]`). Opcional. |

`equipo.json` tampoco se versiona: cada quien tiene el suyo.

### 4. Sustituye el marcador de ruta en los prompts

Los archivos `SKILL.md` y `.claude/commands/*.md` son instrucciones para Claude, no
código, así que no pueden deducir dónde está el proyecto. Llevan el marcador
`<RAIZ_DEL_PROYECTO>`. Sustitúyelo una sola vez por la ruta donde clonaste:

En PowerShell:

```powershell
$raiz = (Get-Location).Path
Get-ChildItem -Recurse -Filter *.md | ForEach-Object {
  (Get-Content $_.FullName -Raw -Encoding utf8).Replace('<RAIZ_DEL_PROYECTO>', $raiz) |
    Set-Content $_.FullName -NoNewline -Encoding utf8
}
```

O en Git Bash:

```bash
grep -rl "<RAIZ_DEL_PROYECTO>" --include=*.md . | xargs sed -i "s|<RAIZ_DEL_PROYECTO>|$(pwd -W)|g"
```

El `-Encoding utf8` no es opcional: Windows escribe en ANSI por defecto y te
rompe los acentos de los prompts.

Esa sustitución deja 11 archivos modificados respecto al repo, y eso chocaría con
el lanzador que actualiza al abrir: `git merge --ff-only` se niega a avanzar si
hay cambios locales. Para que no estorben, se marcan como locales:

```bash
git ls-files -m | xargs -I{} git update-index --skip-worktree "{}"
```

Así el repo conserva el marcador genérico, tu copia tiene la ruta real, y
`git status` queda limpio para que el actualizador funcione. Para revertirlo,
`--no-skip-worktree`. Verlos todos:

```bash
git ls-files -v | grep "^S"
```

### 5. Copia los paquetes de marca

Los informes de churn y los decks de cliente generan HTML con la identidad de GB
Advisors, y esa identidad viene en unos `.skill`, que son ZIP. No están en el
repositorio porque `gb-advisors-design.skill` pesa 12 MB y un binario grande infla
el historial de git para siempre, igual que el `.exe` de GB Print Claude, que se
publica en Releases.

Cópialos a la raíz del proyecto desde la instalación original:

```
gb-advisors-design.skill
gb-deck-cliente.skill
gb-fw-cx-cadence-deck.skill
```

Sin ellos el análisis de churn corre igual, pero el HTML sale sin marca.

### 6. Abre el panel

```
scripts\Panel de Automatizaciones.bat
```

Puedes crear un acceso directo a ese `.bat` en el escritorio.

---

## Los reportes de vTiger que tienes que crear

**Esta es la parte que más cuesta y la que hay que hacer bien.** El panel no consulta
vTiger en vivo para los reportes de métricas: lee los Excel que tú exportas desde el
módulo *Reports* de vTiger a tu carpeta de Descargas. Si un reporte no existe, o le
falta una columna, el script que depende de él falla o calcula de menos.

Crea estos reportes en vTiger, con **exactamente** estos nombres y estas columnas.
El panel los localiza por el patrón del nombre de archivo y toma el más reciente de
cada uno.

### Resumen

| Reporte en vTiger | Archivo que genera | Lo usan |
|---|---|---|
| `Mid Market Metricas 2025 DIQ` | `Mid Market Metricas 2025 DIQ_<fecha>.xlsx` | Calidad CRM, Forecast, Métricas, Compensación, GBS Deal Checker |
| `Mid Market Metricas 2025 - OA` | `Mid Market Metricas 2025 - OA_<fecha>.xlsx` | Churn, Métricas |
| `Mid Market Metricas 2025 - ChurnAssets` | `Mid Market Metricas 2025 - ChurnAssets_<fecha>.xlsx` | Churn, Compensación, Métricas |
| `Mid Market Metricas 2025 - far` | `Mid Market Metricas 2025 - far_<fecha>.xlsx` | Calidad CRM |
| `MIDI Mid Market Metricas 2025 - ORGIQ` | `MIDI Mid Market Metricas 2025 - ORGIQ_<fecha>.xlsx` | Calidad CRM, GBS Organization Maker |
| `Mid Market Metricas 2025 Deals Created` | `Mid Market Metricas 2025 Deals Created_<fecha>.xlsx` | Métricas |
| `Mid Market Metricas 2025 - farmconver` | `Mid Market Metricas 2025 - farmconver_<fecha>.xlsx` | Métricas |

Si prefieres otros nombres, cámbialos en `scripts/generar_reportes_next_step.py`
(constante `CATEGORY_MAP`), pero lo más simple es replicarlos tal cual.

### DIQ (Deals) — 19 columnas

El reporte más importante: casi todo depende de él. Filtro: deals asignados a tu
equipo, sin filtrar por etapa.

```
Deals Organization Name     Products Product Name       Deals Deal ID
Deals Deal Name             Deals Amount                Deals Sales Stage
Deals New/Renew/ProServ     Deals Comm Frequency        Deals Lead Source
Deals Expected Close Date   Organizations GBS Link      Deals Assigned To
Deals GBS Link              Deals Assigned Date         Deals Professional Service Type
Deals Where are we?         Deals Company Segment       Deals Next Step
Deals Modified Time
```

`Deals Modified Time` es imprescindible para las alertas de deals sin atención.
Si falta, el reporte sale pero sin esa sección y el panel avisa.

### OA (Organizaciones y Assets) — 19 columnas

Todos los assets de las organizaciones de tu equipo.

```
Organizations Organization Name   Organizations Organization ID   Organizations Assigned To
Organizations Website             Organizations Employee Count    Organizations Freshworks Tier
Organizations Ranking             Organizations GBS Link          Organizations Billing Country
Organizations Language            Organizations Engagement Level  Organizations NDA
Organizations Created Time        Organizations Industry          Organizations Domain
Assets Asset Name                 Assets Current MRR              Assets Previous MRR
Organizations Last Contacted On
```

### ChurnAssets — 21 columnas

Assets con `Operation = Cancellation` de organizaciones de tu equipo, con fecha de
cambio dentro de los últimos 30 días.

```
Organizations Organization Name   Organizations Organization ID   Organizations Assigned To
Organizations Website             Organizations Employee Count    Organizations Freshworks Tier
Organizations Ranking             Organizations GBS Link          Organizations Billing Country
Organizations Language            Organizations Engagement Level  Organizations Created Time
Organizations Industry            Organizations Domain            Organizations MRR Value
Organizations MRR Result          Assets Asset Name               Assets Operation
Assets VT Change Date             Assets Current MRR              Assets Previous MRR
```

`Assets VT Change Date` es la fecha con la que se acota el churn al trimestre. Sin
ella, un plan de compensación de Q4 cuenta churns de Q3.

### FAR (Farmings) — 25 columnas

Todos los farmings asignados a tu equipo, en cualquier etapa.

```
Organizations Organization Name   Organizations Organization ID   Organizations Assigned To
Organizations Website             Organizations Employee Count    Organizations Freshworks Tier
Organizations Ranking             Organizations GBS Link          Organizations Billing Country
Organizations Language            Organizations Engagement Level  Organizations NDA
Organizations Created Time        Organizations Industry          Organizations Domain
Organizations Last Contacted On   Organizations Account Status    Organizations Account inactive reason
Farming Farming Name              Farming Farming Number          Farming Farming Type
Farming Farming Stage             Farming Assigned To             Farming Next Step
Farming Modified Time
```

### ORGIQ — columnas de Organizations

Todas las organizaciones asignadas a tu equipo. Lleva las columnas de
`Organizations` de la lista de OA; no necesita las de Assets.

### Deals Created — 15 columnas

Deals creados por tu equipo en los últimos 90 días. Mismas columnas que DIQ, menos
`Deal ID`, `Next Step`, `Company Segment` y `Modified Time`, más `Deals Created Time`.

### farmconver — 13 columnas

Deals del equipo que tienen un farming vinculado.

```
Deals Deal Name            Deals Amount                   Deals Sales Stage
Deals Assigned To          Deals Organization Name        Deals Lead Source
Deals New/Renew/ProServ    Deals Professional Service Type  Deals Farming
Organizations Organization Name   Organizations Assigned To
Organizations Billing Country     Organizations Freshworks Tier
```

### Alternativa: bajarlos por API

`scripts/descargar_reportes_vtiger.py` replica estos exports hablando directo con
`webservice.php`, sin pasar por el módulo Reports. Sirve para automatizarlo con el
Programador de tareas.

```bash
python scripts/descargar_reportes_vtiger.py
python scripts/descargar_reportes_vtiger.py --solo orgiq,churn,oa,far
```

Lee la cabecera del archivo antes de confiar en él: ORGIQ, ChurnAssets, OA y FAR
están validados fila por fila contra los exports reales, pero DIQ, Deals Created y
farmconver son aproximaciones cuyo filtro exacto no se pudo inferir. Para
compensación, usa los exports de la UI.

---

## Qué hace cada automatización

El panel agrupa doce tarjetas. Cada una tiene un semáforo que indica si ya se corrió
hoy y un botón para abrir la carpeta de salida.

### Calidad CRM

Audita la higiene del CRM y manda a cada vendedor un correo con lo que le toca
corregir. Seis categorías:

| Alerta | Qué detecta |
|---|---|
| Deals sin Next Step | Deal abierto sin próximo paso definido |
| Deals sin producto | Deal abierto con `Product Name` vacío |
| Deals sin atención | Deal abierto sin modificar en más de 7 días |
| Farmings sin Next Step | Farming activo sin próximo paso |
| Farmings sin producto | Farming activo sin producto |
| Farmings sin atención | Farming activo sin modificar en más de 7 días |

El umbral de 7 días se cambia en `UMBRAL_SIN_ATENCION_DIAS`.

Genera un HTML por vendedor y los adjunta en un correo que se abre en Outlook para
que des el clic final. Entradas: DIQ, FAR, ORGIQ.

### Forecast

Reporte de Upside y Commit por vendedor, en HTML. Permite elegir un rango de meses
libre, no solo el mes en curso. Entrada: DIQ.

### Métricas del trimestre

Métricas ejecutivas con selector de trimestre: el actual hasta la fecha, o cualquiera
de los 8 anteriores ya cerrados. Entradas: DIQ, OA, ChurnAssets, Deals Created,
farmconver.

### Plan de compensación

Un Excel con una hoja por vendedor, para calcular comisiones. Mismo selector de
trimestre. Columnas de cada hoja:

```
Opportunity Name   Deal ID              ECD                 Domain
Comision Rate      MRR                  Payment to GB?      Partner Portal Check
Freshworks ID      Freshwork Invoice    GB INVOICE          Payment date
Amount             Cliente pago         COMM GB             Pendiente para el siguiente Q
```

Las columnas de pago y facturación salen vacías: se llenan a mano. El churn del
trimestre se descuenta acotado por `Assets VT Change Date`, así que un churn de un
trimestre anterior no se cuenta dos veces.

### Análisis de Churn

Lo más elaborado. Cruza ChurnAssets y OA para saber qué organizaciones tienen un
churn **aplicable** (no les queda otro asset activo que lo compense), y sobre cada
una corre un análisis forense de 19 secciones que reconstruye la historia comercial
de la cuenta y emite un veredicto sobre si el churn le es atribuible al vendedor.

Resultado: `resumen-churn-<fecha>.xlsx`, una fila por organización, 15 columnas:

```
Organización        Org ID             Vendedor            Tiempo con la cuenta
Tipo de churn       MRR previo         MRR perdido         ARR perdido
Causa probable      Esfuerzo comercial Veredicto           Confianza
Verificación        Resumen ejecutivo  Informe
```

Los cinco veredictos posibles, con su color en el Excel:

| Veredicto | Color |
|---|---|
| NO APLICA CHURN AL VENDEDOR | Verde |
| NO APLICA, CONTROL O TIEMPO INSUFICIENTE | Verde |
| EVIDENCIA INSUFICIENTE | Amarillo |
| RESPONSABILIDAD COMERCIAL PARCIAL | Naranja |
| APLICA RESPONSABILIDAD DE CHURN | Rojo |

La columna `Informe` enlaza al HTML completo de cada cuenta. **Tarda horas, no
minutos**, porque corre un agente por cuenta. El panel sigue respondiendo mientras
tanto, pero no cierres la ventana.

Cada cuenta analizada guarda su informe, así que una segunda corrida reutiliza lo
que ya existe y solo analiza lo nuevo.

### GBS Deal Checker

Revisa que cada deal abierto tenga su GBS Link bien formado. Salida de 6 columnas,
solo con los que tienen problema:

```
Representante   Deal   Organización   Etapa   Estado GBS Link   Link
```

### GBS Deal Maker y Organization Maker

Generan borradores en Word de deals y organizaciones a crear, para revisarlos antes
de cargarlos.

### Reportes para vendedores, Reasignación de orgs y Rutinas (Sofía)

Automatizaciones alrededor de las reuniones que consigue el equipo de prospección:
reparto rotativo de presentaciones entre vendedores, detección de menciones sobre
organizaciones con señal de deal para reasignarlas, y preparación de los decks de
cliente.

### Validar creación de deals

Revisa que los deals recién creados cumplan las reglas mínimas de campos.

### Cadence Generator

Deck HTML del pipeline para la cadencia con el fabricante, con notas editables que
se guardan en el propio documento.

---

## Rutinas desatendidas

`Rutinas/Run-Rutina.ps1` corre una automatización sin supervisión, con un marcador
diario para no repetirla:

```powershell
.\Rutinas\Run-Rutina.ps1 -Skill daily-client-presentation-prep
```

Para probar sin disparar nada, informando qué haría:

```powershell
.\Rutinas\Run-Rutina.ps1 -Skill daily-client-presentation-prep -DryRun
```

`Rutinas/Registrar-Tareas.ps1` las registra en el Programador de tareas de Windows.

---

## Estructura

```
scripts/                    Generadores y el panel. Los .bat son sus lanzadores.
  paths.py                  Resuelve todas las rutas desde su propia ubicación.
  equipo.py                 Carga equipo.json. Punto único del equipo.
  load_env.py               Carga .env.
  mail_helper.py            Habla con Outlook por COM.
  brand/                    Tipografías y logos para los HTML.
Claude/Scheduled/           Prompts de las automatizaciones con IA (SKILL.md).
.claude/commands/           Comandos de proyecto para Claude Code.
Churn Accounts/.claude/     Skill del análisis forense de churn.
Rutinas/                    Corredor y registro de tareas programadas.
```

Las carpetas de salida (`Forecast/`, `Metricas Ejecutivas/`, `Churn Accounts/informes/`)
se crean solas y no se versionan: se regeneran desde vTiger.

---

## Si algo falla

| Síntoma | Causa y arreglo |
|---|---|
| `falta equipo.json` | No copiaste `equipo.example.json`. Ver paso 3. |
| Un vendedor no aparece en ningún reporte | Su `nombre` en `equipo.json` no coincide exacto con `Assigned To` de vTiger. Revisa acentos y espacios. |
| `ERROR [StopIteration]` | El export de vTiger vino sin filas. Normal al inicio de un trimestre. |
| `PermissionError` al guardar un Excel | Lo tienes abierto en Excel. Los scripts guardan al lado con sufijo `-2` y avisan. |
| `ACCESS_DENIED` de vTiger | Casi siempre el campo no existe en ese módulo, no es un problema de permisos. |
| Acentos rotos en una salida | Windows escribe en ANSI por defecto. Los scripts corren con `PYTHONIOENCODING=utf-8`; desde PowerShell pasa `-Encoding utf8` explícito. |
| El panel no arranca | Falta un paquete. `pip install openpyxl pypdf pillow pywin32 PySide6`. |

---

## Notas sobre vTiger

VQL no es SQL. `LIMIT <offset>, <count>` lleva el offset primero, no hay `JOIN` ni
`GROUP BY`, y los vacíos no se filtran con `WHERE campo = ''`. Una consulta devuelve
como máximo 200 filas, así que los barridos se paginan.

---

## Seguridad

- La access key va solo en `.env`, nunca en el código ni en los prompts.
- `.env` y `equipo.json` están en `.gitignore`.
- Los informes de churn contienen juicios sobre el desempeño de personas con nombre
  y apellido, y los planes de compensación llevan cifras de comisión. Esas carpetas
  están excluidas del repositorio a propósito. Si cambias el `.gitignore`, ten claro
  qué estás publicando.
