# `_dev/` - Carpeta de trabajo del proyecto

Carpeta para herramientas de desarrollo, snapshots y referencias. **Toda la
logica viva esta en `robot.db`**; lo de aqui es soporte para no quemar tokens
ni tiempo en cada sesion.

## Estructura

```
_dev/
  README.md              <-- este archivo
  conventions.md         <-- reglas de Rocketbot, patrones, plantillas, paths
  lib/
    dbtools.py           <-- decoder/encoder de robot.db (CLI + libreria)
    autodetect_prologue.py  <-- prologo de ~95 lineas que va al inicio de cada bot hijo
  snapshot/
    generate.py          <-- regenera los snapshots desde robot.db
    bot_NN_<nombre>/
      meta.json          <-- vars del modulo + arbol de comandos sin scripts
      cmd_XX.py          <-- script Python de cada execScriptPython
      cmd_07_children__04.py  <-- nested en evaluateIf usando . como separador
```

## Workflows comunes

### Inspeccionar un bot sin gastar tokens leyendo robot.db

```
# CLI
python _dev/lib/dbtools.py list                # lista bots
python _dev/lib/dbtools.py vars 13             # vars del modulo
python _dev/lib/dbtools.py cmds 32             # comandos (top-level)
python _dev/lib/dbtools.py dump 13 0           # script Python al stdout
python _dev/lib/dbtools.py grep RIVERSIDE -i   # busca en todos los scripts
```

O simplemente leer los archivos en `_dev/snapshot/bot_NN_<nombre>/`.

### Regenerar el snapshot tras cambios

```
python _dev/snapshot/generate.py
```

Hacer **siempre** despues de aplicar un parche, antes de cerrar la sesion, asi
el snapshot del repo refleja el estado de robot.db.

### Modificar un script de un bot

Opcion A (programatica, con verificacion de compilacion):

```python
import sys
sys.path.insert(0, r'c:\Repositorios\MPF - Rocketbot\Rocketbot---MPF\_dev\lib')
from dbtools import RobotDB
db = RobotDB()
script = db.get_command(13, 0)['command']
# ... modificar script ...
db.update_command_script(13, 0, script)   # compila y guarda
```

Opcion B (parche one-shot): script en `.audit/patch_*.py` (gitignored).

### Verificar conformidad con las convenciones

`conventions.md` lista todo lo que el codigo asume. Antes de escribir un
parche grande, leer:

- Seccion 1: las 6 reglas duras de Rocketbot
- Seccion 5: estructura del Registro de Coordinacion
- Seccion 7: plantillas de correo y idiomas
- Seccion 9: variables del modulo requeridas

## Notas de mantenimiento

- `_dev/snapshot/` es generado, **no editar a mano**. Si hay que cambiar
  algo, modificar el bot via dbtools y regenerar.
- `_dev/lib/dbtools.py` y `autodetect_prologue.py` son las unicas fuentes
  de verdad sobre como manipular la DB. Si cambian patrones del prologo
  (paths, credenciales) actualizar aqui Y replicar al snapshot.
- Las plantillas de correo del codigo viven embebidas en los scripts de
  bot 13 y bot 12. Si hay que duplicarlas o moverlas, considerar primero
  factorizar en `_dev/lib/email_templates.py`.
