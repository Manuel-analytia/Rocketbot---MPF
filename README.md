# Rocketbot — MPF (SAN ISIDRO)

RPA en Rocketbot para el flujo semanal COMEX de MPF con el cliente **RIVERSIDE NATURAL FOODS, LTD**.

## Qué hace

Un orquestador y 5 hijos automatizan:

| Bot | Función |
|---|---|
| **32** `0_Flujo_Regular` | Orquestador (no contiene lógica de negocio) |
| **29** `0_1_Verificacion_Condiciones` | Precheck: rutas, Config_Rutas.xlsx, Registro anual |
| **26** `1_Crear_Carpetas` | `{ROOT_EMBARQUES}\{AÑO}\Semana N` (y N+1 si es domingo) |
| **30** `2_Embarques` | Lee Excel semanal, filtra RIVERSIDE, registra nuevos |
| **13** `3_Revision_COAS_Informe` | Consolida PDF COAs+Informes y envía al cliente |
| **12** `4_Revision_Invoice` | Envía Invoice cuando aparece en `\Proveedores MPF - RIVERSIDE...` |

## Estructura del repo

```
Rocketbot---MPF/
├── robot.db                # SQLite con todos los bots (base64/JSON)
├── README.md               # este archivo
├── Funcional/
│   ├── Doc. Funcional corregido.docx        # spec del cliente
│   ├── Guia de Trabajo Rocketbot.txt        # 10 reglas duras + convenciones
│   ├── Estado y Pendientes.txt              # estado actual del proyecto
│   ├── Ejemplo de archivo de embarque.xlsx
│   ├── Ejemplo de archivo COAS.xlsx
│   ├── bot_comex.bat                        # lanzador del server clásico
│   └── bot_comex_from_repo.bat              # lanzador que copia desde OneDrive
└── .audit/                 # patches reproducibles (gitignored)
    ├── patch_*.py          # 20 activos (cambios actuales en robot.db)
    ├── backups/            # backup .b64 por cada patch aplicado
    └── legacy/             # 67 patches superseded
```

## Estructura de datos del cliente (OneDrive)

```
{OneDrive}\
├── SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES\
│   ├── 2026\Semana N\Semana N.xlsx       # input semanal del cliente
│   ├── COORDINACION DE EMBARQUES\
│   │   └── Registro de COAS e informes completos {AÑO}.xlsx
│   └── 0. Bot\                            # carpeta técnica del bot
│       ├── robot.db                       # copia distribuida del bot
│       ├── Logs\LOG - YYYY-MM-DD HH.MM.SS.txt
│       ├── Rutas\Config_Rutas.xlsx        # parámetros del cliente
│       ├── Correos\{Invoice,COAS_e_Informes,Estatus}\{Para,CC,CO}.txt
│       ├── Pruebas\pruebas.txt            # forzar semana (N|YYYY)
│       └── Temp\                          # PDFs consolidados temporales
├── Proveedores MPF - RIVERSIDE NATURAL FOODS, LTD\
│   └── {embarque} RIVERSIDE ...\          # invoice del cliente + consolidado
├── Archivos de Juliana Legua - ORGANICO\  # fuente PDFs de COAs
└── PISCO - {AÑO}\                         # fuente PDFs de Informes
```

## Flujo de despliegue

```
Local (dev)                OneDrive (distribución)            Servidor MPFBOOT01
─────────────              ──────────────────────             ──────────────────
Aplicar patch              ⬇️ sincroniza                       bot_comex_from_repo.bat:
  python .audit/...   ──▶  0. Bot\robot.db          ──▶        copy robot.db
Validar (ast/test)                                              rocketbot.exe -start=...
Commit en repo
Subir a OneDrive
```

## Cómo aplicar un cambio

1. **Diagnóstico** — qué hace hoy, qué falla.
2. **Patch** — crear `.audit/patch_*.py` con `OLD`/`NEW`, idempotente por marker, con backup auto.
3. **Aplicar** — `python .audit/patch_<algo>.py`.
4. **Validar** — `ast.parse` del comando; tests sandbox cuando aplique.
5. **Distribuir** — copiar `robot.db` a `OneDrive\...\0. Bot\robot.db`.
6. **REGLA 9** — en el servidor, cerrar Rocketbot completo antes de la próxima corrida (el `.bat` lo hace).

Ver `Funcional/Guia de Trabajo Rocketbot.txt` para las **10 reglas duras** del entorno.
Ver `Funcional/Estado y Pendientes.txt` para el estado actual y patches aplicados.
