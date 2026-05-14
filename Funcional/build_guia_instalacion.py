"""
Genera 'Funcional/Guia de Instalacion Rapida.docx' usando la plantilla
corporativa de Funcional/_template/.

COMPORTAMIENTO:
  - Si el .docx YA EXISTE: carga el archivo, borra todo desde la primera H1
    hacia abajo y reescribe el body. Esto preserva la portada y la tabla de
    contenido que tu hayas insertado manualmente en Word.
  - Si NO EXISTE: lo crea desde cero con la portada. La tabla de contenido
    debes insertarla una unica vez via Word: Referencias > Tabla de contenido.

Para regenerar despues de editar contenido:
    python Funcional/build_guia_instalacion.py
"""
from __future__ import annotations
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / '_template'))
from build_doc import DocumentBuilder  # noqa: E402


def main() -> None:
    out = HERE / 'Guia de Instalacion Rapida.docx'

    if out.exists():
        # Preserva portada + TOC manual del usuario
        db = DocumentBuilder.from_existing(out)
    else:
        # Primera generación: portada nueva. La TOC la inserta el usuario en Word.
        db = DocumentBuilder()
        db.add_cover(
            title='Guía de Instalación Rápida',
            subtitle='Bot RPA - Envío de correos COMEX',
            date='Mayo 2026',
        )

    # ============ 1. INTRODUCCIÓN ============
    db.h1('1. Introducción')
    db.p('Bot RPA que automatiza el envío de Certificados de Calidad e Informes '
         'de Ensayo (COAs) consolidados, así como el envío de Invoices, a los '
         'destinatarios del cliente Riverside Natural Foods, Ltd. Corre '
         'desatendido todos los días en el servidor, bajo la cuenta estándar '
         'del bot, ejecutado por el Programador de Tareas de Windows.')
    db.p('La cuenta del bot ya tiene OneDrive sincronizado con las cuatro '
         'carpetas operativas (SAN ISIDRO Embarques, SAN ISIDRO Riverside, '
         'Archivos de Juliana Legua ORGÁNICO y PISCO {AÑO}). El archivo '
         'Config_Rutas.xlsx, las listas de destinatarios y el Registro de '
         'Coordinación del año en curso ya están configurados; TI no debe '
         'modificarlos.')

    # ============ 2. INSTALACIÓN ============
    db.h1('2. Instalación')

    db.h2('2.1 Importar el bot a Rocketbot Studio')
    db.p('Requisito previo: la cuenta del bot debe tener una Licencia de Producción '
         'activa en esta instalación de Rocketbot. Sin ella, el CLI -start= no '
         'funciona y la corrida queda bloqueada con el mensaje "Does not have a '
         'production license" en consola.')
    db.p('Instalar Rocketbot Studio versión 2025 en el servidor. Antes de '
         'abrirlo por primera vez, copiar el archivo robot.db entregado por el '
         'equipo de RPA a la carpeta de instalación de Rocketbot (por defecto '
         'C:\\Program Files (x86)\\Rocketbot\\robot.db), reemplazando el archivo '
         'que viene con la instalación.')
    db.p('Abrir Rocketbot Studio una primera vez con permisos de administrador. '
         'Si aparece el aviso de actualización de drivers, aceptar y dejar que '
         'termine; esto evita que el popup interrumpa las corridas posteriores '
         'desatendidas. Luego activar la Licencia de Producción si aún no está '
         'aplicada en esta máquina.')
    db.p('Verificar en Rocketbot Studio que aparezcan los seis robots del '
         'proyecto: 0_Flujo_Regular (orquestador), 0_1_Verificacion_Condiciones, '
         '1_Crear_Carpetas, 2_Embarques, 3_Revision_COAS_Informe y '
         '4_Revision_Invoice.')

    db.h2('2.2 Colocar el archivo .bat')
    db.p('Crear la carpeta C:\\Scripts si no existe, y colocar el archivo '
         'bot_comex.bat (entregado junto con el robot.db) en esa ruta. Esa '
         'será la ruta que ejecutará el Programador de Tareas. La estructura '
         'del .bat es la siguiente:')
    db.code('@echo off')
    db.code('REM 1. Asegurar que OneDrive esté corriendo')
    db.code('start "" "%LOCALAPPDATA%\\Microsoft\\OneDrive\\OneDrive.exe" /background')
    db.code('')
    db.code('REM 2. Pausa para que OneDrive sincronice')
    db.code('timeout /t 120 /nobreak')
    db.code('')
    db.code('REM 3. Ejecutar el orquestador en Rocketbot')
    db.code('cd /d "C:\\Program Files (x86)\\Rocketbot"')
    db.code('start "" /WAIT "rocketbot.exe" -start=0_Flujo_Regular')
    db.code('exit /b 0')
    db.p('Si la red del servidor es lenta, subir el timeout de 120 a 240 '
         'segundos para que OneDrive tenga tiempo de sincronizar el Excel '
         'semanal antes de invocar al orquestador.')

    db.h2('2.3 Crear la tarea programada')
    db.p('Desde el Programador de Tareas de Windows (taskschd.msc), crear una '
         'tarea nueva con la siguiente configuración:')
    db.table(
        headers=['Campo', 'Valor'],
        rows=[
            ['Nombre',                       'Bot COMEX MPF'],
            ['Cuenta de ejecución',          'Cuenta del bot (la misma que sincroniza OneDrive)'],
            ['Privilegios',                  'Desmarcar "Ejecutar con privilegios más altos" (recomendación oficial de Rocketbot ante errores de módulo o licencia)'],
            ['Cuándo ejecutar',              'Tanto si el usuario inició sesión como si no'],
            ['Desencadenador',               'Diario, repetir cada un día'],
            ['Hora de inicio',               '07:00 (sugerencia; ajustar según necesidad)'],
            ['Acción - Programa',            'C:\\Scripts\\bot_comex.bat'],
            ['Acción - Iniciar en',          'C:\\Program Files (x86)\\Rocketbot  (importante: sin esto Rocketbot falla al ubicar sus recursos)'],
            ['Argumentos',                   '(ninguno)'],
            ['Si el equipo no está en CA',   'Permitir ejecutar igualmente'],
            ['Si la tarea se atrasa',        'Ejecutar cuanto antes'],
            ['Tiempo máximo',                'Detener la tarea si dura más de 4 horas'],
        ],
        col_widths_in=[2.0, 4.0],
    )

    # ============ 3. VERIFICACIÓN ============
    db.h1('3. Verificación')

    db.h2('3.1 Smoke test manual con semana antigua')
    db.p('Editar el archivo de pruebas:')
    db.code('<OneDrive>\\SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES\\0. Bot\\Pruebas\\pruebas.txt')
    db.p('Colocar dentro una sola línea con el formato semana|año, por ejemplo:')
    db.code('7|2026')
    db.p('Esto fuerza al bot a procesar la semana ISO 7 del año 2026 (que ya '
         'tiene datos cargados de pruebas previas) en lugar de la semana '
         'actual. Guardar el archivo y ejecutar el .bat manualmente desde una '
         'terminal:')
    db.code('C:\\Scripts\\bot_comex.bat')
    db.p('Esperar a que termine y revisar el log más reciente:')
    db.code('<OneDrive>\\SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES\\0. Bot\\Logs\\LOG - YYYY-MM-DD HH.MM.SS.txt')
    db.p('El log debe terminar con la línea de cierre:')
    db.code('Resumen orquestador: FINALIZADO | PRECHECK=OK | H1=EJECUTADO | H2=OK | H3=EJECUTADO | H4=EJECUTADO')
    db.p('Una vez validado, vaciar el archivo pruebas.txt (sin contenido o '
         'solo con un comentario que empiece con #). Si queda con un número '
         'adentro, el bot procesará siempre esa semana en lugar de la actual.')

    db.h2('3.2 Smoke test desatendido')
    db.p('Cerrar la sesión del servidor sin apagarlo. Esperar a la hora del '
         'trigger configurado (o adelantarlo unos minutos para no esperar 24 '
         'horas). Volver a iniciar sesión y verificar que existe un nuevo '
         'archivo LOG con la fecha y hora del trigger, que termina con la '
         'misma línea de FINALIZADO de la prueba anterior, y que las carpetas '
         'y registros se actualizaron en OneDrive.')

    # ============ 4. TROUBLESHOOTING ============
    db.h1('4. Troubleshooting')
    db.p('Diagnóstico rápido según el síntoma observado en el log:')
    db.table(
        headers=['Síntoma', 'Causa probable', 'Qué revisar'],
        rows=[
            ['Consola dice "Does not have a production license"',
             'Licencia de Producción no activada en esta máquina',
             'Activar la licencia desde Rocketbot Studio (Menú > Licencia) con la cuenta del bot'],
            ['La tarea no se ejecuta',
             'Permisos de la cuenta del bot, o "Iniciar en" mal configurado',
             'secpol.msc > Asignación de derechos > "Iniciar sesión como tarea por lotes". Y verificar que "Iniciar en" apunte a la carpeta de Rocketbot'],
            ['Log dice RUTA_SEMANA_INVALIDA o ERROR_NOT_VAR',
             'OneDrive desincronizado',
             'Marcar las cuatro carpetas como "Disponibles en este dispositivo"'],
            ['Log dice "Falla envío correo: ..."',
             'SMTP_AUTH deshabilitado en O365 o credencial vencida',
             'El detalle indica el paso (connect / login / sendmail). Coordinar con admin de O365 para habilitar SMTP AUTH en el buzón del bot'],
            ['Log dice SIN_DATA o no detecta el Excel',
             'OneDrive no terminó de sincronizar el Excel semanal',
             'Subir el timeout del .bat de 120 a 240 segundos, o forzar sincronización manual antes de ejecutar'],
            ['El log se corta a medias',
             'La tarea se detuvo por timeout (4 h)',
             'Revisar el log central; lo más común es un archivo PDF muy grande que tarda en subir por SMTP (ver fila SMTP)'],
        ],
        col_widths_in=[1.8, 1.8, 2.6],
    )
    db.p('Dónde mirar primero ante cualquier duda: la carpeta Logs en 0. Bot '
         'es la fuente única de verdad de cada corrida. Cada línea identifica '
         'el bot que la generó (3_Revision_COAS_Informe, 4_Revision_Invoice, '
         'etc.) e incluye un resumen al final por bot.')

    # ============ GUARDAR ============
    db.save(out)
    print(f'Generado: {out}')


if __name__ == '__main__':
    main()
