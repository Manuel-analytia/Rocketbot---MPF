import os
from datetime import datetime
def _log(proceso, msg):
    rp = GetVar('v_orq_log_path')
    if rp and os.path.isabs(rp):
        try:
            with open(rp, 'a', encoding='utf-8') as _f:
                _f.write(f"{datetime.now().strftime('%H:%M:%S')}|{proceso}|{msg}\n")
        except Exception:
            pass
ruta_anio    = GetVar('v_ruta_anio') or ''
nombre_sem   = GetVar('v_nombre_semana') or ''
ruta_semana  = os.path.join(ruta_anio, nombre_sem) if ruta_anio and nombre_sem else ''
SetVar('v_ruta_semana', ruta_semana)
_log('1_Crear_Carpetas', f'DEBUG rutas | v_ruta_anio={ruta_anio!r} | v_nombre_semana={nombre_sem!r} | v_ruta_semana={ruta_semana!r}')
