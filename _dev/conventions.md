# Convenciones del proyecto Bot RPA MPF (SAN ISIDRO)

Documento de referencia rapida. **Si algo aqui contradice el codigo, el codigo
gana** — al modificar codigo, actualizar este documento. Para ver el estado
actual de un bot, leer su carpeta en `_dev/snapshot/`.

---

## 1. Reglas duras de Rocketbot

| # | Regla | Por que |
|---|-------|---------|
| 1 | `setVar` con `GetVar(...)` deja la var en `'ERROR_NOT_VAR'`. Usar siempre `execScriptPython` con `SetVar('v_x', os.path.join(GetVar('a'), GetVar('b')))`. | Rocketbot no evalua `GetVar()` en expresiones de `setVar`. |
| 2 | Variables entre bots hijos **NO** se heredan. Cada `execRocketBotDB` crea un scope nuevo donde `GetVar` de variables ajenas devuelve `'ERROR_NOT_VAR'`. | Rocketbot aisla scopes por bot. |
| 3 | Variables **no declaradas en el modulo del bot** se reinician a `'ERROR_NOT_VAR'` entre cada `execScriptPython`. Solo las del modulo persisten. | Comportamiento documentado por experimentacion. |
| 4 | `GetVar` de undefined retorna literal `'ERROR_NOT_VAR'` (no `''` ni `None`). Siempre validar `if v and v != 'ERROR_NOT_VAR':`. | |
| 5 | `execRocketBotDB` resuelve sub-flujo por **nombre** (no por id). Si hay varios bots con el mismo nombre, Rocketbot elige uno (en este proyecto solo hay 1 de cada). | |
| 6 | Funciones helper Python definidas en un script **no ven** `GetVar`/`SetVar` desde dentro de OTRA funcion helper. Para escribir al log desde una funcion helper hay que reproducir el bloque inline (ver `_log` y `_log_line` en bots 13/30). | Confirmar caso por caso. |

## 2. Arquitectura del flujo (orquestador 0_Flujo_Regular = bot 32)

```
bot 32 (0_Flujo_Regular)
  CMD03 -> bot 29 (0_1_Verificacion_Condiciones)        Precheck + crea Registro
  CMD08 -> bot 26 (1_Crear_Carpetas)                    Crea carpeta semana
  CMD11 -> bot 30 (2_Embarques)                         Lee Excel, escribe Registro
  CMD12 evaluateIf hay data:
    -> bot 13 (3_Revision_COAS_Informe)                 Consolida y envia COAs+Informes
    -> bot 12 (4_Revision_Invoice)                      Envia Invoice
```

Cadencia: externa via Programador de Tareas de Windows. No hay cron interno.

## 3. Paths y nomenclatura

| Identificador | Patron |
|---------------|--------|
| OneDrive root | `C:\Users\{USERNAME}\OneDrive - mpf.com.pe` (autodetect) |
| ROOT_EMBARQUES | `{OneDrive}\SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES` |
| ROOT_RIVERSIDE | `{OneDrive}\SAN ISIDRO - RIVERSIDE NATURAL FOODS, LTD` |
| ROOT_COAS_FUENTE | `{OneDrive}\Archivos de Juliana Legua - ORGANICO` |
| ROOT_INFORMES | `{OneDrive}\PISCO - {AÑO_ISO}` (prefijo `"PISCO - "` configurable) |
| BOT_ROOT | `{ROOT_EMBARQUES}\0. Bot` |
| Carpeta semanal | `{ROOT_EMBARQUES}\{YEAR}\Semana {N}` |
| Carpeta embarque en RS | `{ROOT_RIVERSIDE}\<EMB> RIVERSIDE <MMM> <YY>` ej: `2000009376 RIVERSIDE MAR 26` |
| PDF consolidado | `<EMB> - Quality Certificate & Test Reports.pdf` |
| Registro Coordinacion | `{ROOT_EMBARQUES}\COORDINACION DE EMBARQUES\Registro de COAS e informes completos {AÑO}.xlsx` |
| Log | `{BOT_ROOT}\Logs\LOG - YYYY-MM-DD HH.MM.SS.txt` |
| Envios.xlsx | `{BOT_ROOT}\Logs\Envios.xlsx` |
| pruebas.txt | `{BOT_ROOT}\Pruebas\pruebas.txt`. Formato: una linea, `N` (semana) o `N|YYYY` (semana + año). Vacia = semana ISO actual. |

## 4. Config_Rutas.xlsx

Ubicacion: `{BOT_ROOT}\Rutas\Config_Rutas.xlsx`. Hojas:

**BASE**
| PARAMETRO | VALOR DEFAULT |
|-----------|---------------|
| ROOT_EMBARQUES | `SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES` |
| ROOT_RIVERSIDE | `SAN ISIDRO - RIVERSIDE NATURAL FOODS, LTD` |
| ROOT_COAS_FUENTE | `Archivos de Juliana Legua - ORGANICO` |
| ROOT_INFORMES_PREFIJO | `PISCO - ` (espacio final, el bot concatena el año ISO) |

**SUBRUTAS_BOT**: `BOT_ROOT, BOT_LOGS, BOT_RUTAS, BOT_CORREOS, BOT_PRUEBAS, BOT_TEMP, BOT_EVIDENCIAS`.

**SUBRUTAS_NEGOCIO**
| PARAMETRO | VALOR DEFAULT |
|-----------|---------------|
| CARPETA_COORDINACION_EMBARQUES | `COORDINACION DE EMBARQUES` |
| NOMBRE_REGISTRO_PREFIJO | `Registro de COAS e informes completos ` |
| PAGINA_EXCLUIR_INFORMES | `TERMINOS Y CONDICIONES DE LOS SERVICIOS` |

**CREDENCIALES**: `SMTP_HOST, SMTP_PORT, SMTP_SSL, SMTP_USER, SMTP_PASS`.

Eliminadas (legacy): `URL_*`, `MICROSOFT_*`, `IMAP_*`.

## 5. Registro de Coordinacion - estructura del .xlsx

12 hojas: `ENE FEB MAR ABR MAY JUN JUL AGO SEP OCT NOV DIC`. La hoja de cada
embarque se determina por el mes del **lunes ISO** de su semana.

| Col | Letra | Header | Quien escribe |
|-----|-------|--------|---------------|
| 1 | A | (vacia, originalmente Semana) | bot 30 |
| 2 | B | Cliente | bot 30 |
| 3 | C | Embarque | bot 30 |
| 4 | D | Fecha | bot 30 |
| 5 | E | N de pedido | bot 30 |
| 6 | F | COAS | **cliente manual** ("COMPLETO") |
| 7 | G | Informes | **cliente manual** ("COMPLETO") |
| 8 | H | Estado Invoice | bot 12 ("ENVIADO") |
| 9 | I | Fecha de envio Invoice | bot 12 (fecha hoy) |
| 10 | J | Existencia | bot 13 (`OK / COAS_NO_ENCONTRADO / INFORMES_NO_ENCONTRADO / AMBOS_NO_ENCONTRADOS`) |
| 11 | K | Fecha de envio COAs | bot 13 (timestamp Lima) |

Bot 30 escribe **solo B-E**. Bot 13 lee F (COAS) y G (Informes), escribe J y K.
Bot 12 lee D (Fecha) y H (Estado), escribe H e I.

## 6. Reglas por bot

### Bot 13 (3_Revision_COAS_Informe)
- Procesa filas con `K` vacia (no enviadas) y `F=COMPLETO AND G=COMPLETO`.
- Si F o G no son COMPLETO **y** dias_atraso > **4** -> agrega a alerta HTML a `Correos/Estatus/`.
- Doble alerta (color rojo) si dias > **8**.
- Busca COAs en `{ROOT_COAS_FUENTE}` (archivo plano cuyo nombre contenga el embarque).
- Busca Informes en `{ROOT_INFORMES}/{embarque}/` (subcarpeta).
- Recorta paginas T&C usando texto de `PAGINA_EXCLUIR_INFORMES`.
- Consolida con PyPDF3/pypdf/PyPDF2/fitz (en orden de preferencia).
- **Copia a Riverside SIEMPRE** (antes del envio del correo, para no perder el archivo si SMTP falla).
- Patron de carpeta destino: `<EMB> RIVERSIDE <MMM> <YY>`.
- Antes de consolidar, **borra** copia previa en `{BOT_TEMP}/`.
- Plantilla envio: **INGLES**. Asunto `PO-{ped} || OP-{emb} || Riverside`.
- Plantilla faltantes: ESPAÑOL, va a `Correos/Estatus/`.
- Plantilla alerta no-completos: ESPAÑOL HTML, va a `Correos/Estatus/`.

### Bot 12 (4_Revision_Invoice)
- Procesa filas con `I` vacia. Logica por fecha:
  - `fecha_emb > hoy` -> futuro, skip
  - `dias = hoy - fecha_emb < 1` -> mismo dia, skip
  - `dias > 1` -> log `WARNING INVOICE_ATRASADO`
- Busca carpeta del embarque en `{ROOT_RIVERSIDE}` por prefijo
  (match exacto o `<EMB> ...`).
- Busca archivos cuyo nombre contenga `invoice` (case insensitive).
- 1 correo por archivo encontrado (no consolida).
- Plantilla envio: **INGLES**. Asunto: si nombre tiene `invoice po <X>` ->
  `PO-<X> / OP-{emb}`, sino `PO-{ped} / OP-{emb}`.
- Plantilla pendientes: ESPAÑOL HTML, va a `Correos/Estatus/`.

### Bot 30 (2_Embarques)
- Lee TODOS los `.xlsx` en la carpeta semanal (no solo el mas reciente).
- Filtra columna 6 (cliente) por `RIVERSIDE NATURAL FOODS, LTD.` (case insensitive substring).
- Lee: emb=col 0, fecha=col 3, pedido=col 8.
- Dedup cross-files por `embarque|pedido`.
- Calcula hoja del Registro = mes del lunes ISO.
- Escribe filas nuevas (no presentes) en B-E.
- Si hay nuevos -> notifica a `Correos/COAS_e_Informes/` (plantilla ESPAÑOL).

### Bot 26 (1_Crear_Carpetas)
- Crea **solo** la semana actual: `{ROOT_EMBARQUES}/{AÑO}/Semana {N}/`.
- Detecta el Excel mas reciente en esa carpeta y setea `v_ruta_excel`.

### Bot 29 (0_1_Verificacion_Condiciones / Precheck)
- Crea `{BOT_ROOT}` y subcarpetas (Logs, Rutas, Correos/{COAS_e_Informes,Invoice,Estatus}, Pruebas, Temp).
- Crea `Config_Rutas.xlsx` con template si no existe.
- Crea **Registro Coordinacion del año actual** si no existe (12 hojas vacias con headers).
- Lee `pruebas.txt` para forzar semana/año.
- Lee credenciales SMTP de Config_Rutas y las setea como variables.
- Setea `v_orq_log_path` para que los hijos escriban al log central.

### Bot 32 (0_Flujo_Regular)
- Lo unico que hace: orquestar. No hace logica de negocio.
- Logging detallado: `Inicio del flujo`, `Iniciando X`, `X EJECUTADO | detalle`, `OMITIDO`.

## 7. Plantillas de correo

| Bot | Tipo | Idioma | Destinatarios |
|-----|------|--------|---------------|
| 13 | Envio consolidado COAs+Informes | INGLES | `Correos/COAS_e_Informes/` |
| 13 | Faltantes (archivo no encontrado) | ESPAÑOL | `Correos/Estatus/` |
| 13 | Alerta HTML no-completos > 4 dias | ESPAÑOL | `Correos/Estatus/` |
| 12 | Envio Invoice | INGLES | `Correos/Invoice/` |
| 12 | Pendientes HTML | ESPAÑOL | `Correos/Estatus/` |
| 30 | Embarques nuevos | ESPAÑOL | `Correos/COAS_e_Informes/` |

Asunto envio COAs: `PO-{pedido} || OP-{embarque} || Riverside` (separadores `||`).
Asunto envio Invoice: `PO-{pedido_o_po_extraido} / OP-{embarque}` (separador `/`).
Firma comun (EN): `Best regards,\nCOMEX MPF Bot`.

## 8. SMTP

- Host: `smtp.office365.com:587` STARTTLS.
- Timeout actual: **300s** (subido para tolerar runtime de Rocketbot lento con adjuntos grandes).
- `_enviar_correo` retorna `(ok: bool, err: str)`. El caller incluye `err` en el log de fallo.
- Stages que el `err` reporta: `build_msg`, `connect ...`, `starttls`, `login ...`, `sendmail to N dest (adjunto X.XX MB)`.

## 9. Variables del modulo (las criticas)

Cada bot debe declarar en su modulo todas las variables que **tienen que persistir**
entre comandos (regla 3). Lista minima requerida en bots hijos:

- `v_root_onedrive`, `v_root_embarques`, `v_root_riverside`, `v_root_coas_fuente`, `v_root_informes_prefijo`
- `v_bot_root`, `v_bot_logs`, `v_bot_temp`, `v_bot_rutas`, `v_bot_correos`, `v_bot_pruebas`
- `v_ruta_correos_coas`, `v_ruta_correos_invoice`, `v_ruta_correos_estatus`
- `v_anio_semana`, `v_numero_semana_trabajo`, `v_nombre_semana`, `v_mes_semana`
- `v_ruta_anio`, `v_ruta_semana`, `v_ruta_excel`, `v_ruta_registro_coordinacion`
- `v_smtp_host`, `v_smtp_port`, `v_smtp_user`, `v_smtp_pass`, `v_smtp_ssl`
- `v_orq_log_path`, `v_ruta_log_actual`
- Estado del bot: `v_h{N}_estado`, `v_h{N}_detalle`

El AUTODETECT_PROLOGUE inyectado al inicio de cada script setea todas estas
desde cero (lee USERNAME, lee Config_Rutas.xlsx, lee pruebas.txt). Ver
[`_dev/lib/autodetect_prologue.py`](lib/autodetect_prologue.py).
