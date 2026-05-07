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
ruta_base    = GetVar('v_ruta_base') or ''
anio_semana  = GetVar('v_anio_semana') or ''
ruta_anio    = os.path.join(ruta_base, anio_semana) if ruta_base and anio_semana else ''
SetVar('v_ruta_anio', ruta_anio)
_log('1_Crear_Carpetas', f'DEBUG rutas | v_ruta_base={ruta_base!r} | v_anio_semana={anio_semana!r} | v_ruta_anio={ruta_anio!r}')
