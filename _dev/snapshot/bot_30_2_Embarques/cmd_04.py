import os
from datetime import datetime

def _log_safe(proceso, msg):
    for varname in ('v_orq_log_path', 'v_ruta_log_actual'):
        rp = GetVar(varname)
        if rp and os.path.isabs(rp):
            try:
                with open(rp, 'a', encoding='utf-8') as _f:
                    _f.write(f"{datetime.now().strftime('%H:%M:%S')}|{proceso}|{msg}\n")
                return
            except Exception:
                pass

ruta_base   = GetVar('v_ruta_base') or ''
anio_semana = GetVar('v_anio_semana') or ''
ruta_anio   = GetVar('v_ruta_anio') or ''
nombre_sem  = GetVar('v_nombre_semana') or ''
ruta_semana = GetVar('v_ruta_semana') or ''

_log_safe('2_Embarques', f'DEBUG rutas | base={ruta_base!r} | anio={anio_semana!r} | ruta_anio={ruta_anio!r} | nombre_sem={nombre_sem!r} | ruta_semana={ruta_semana!r}')

todos_archivos = []
archivos_excel = []
if ruta_semana and os.path.isdir(ruta_semana):
    try:
        for nombre in sorted(os.listdir(ruta_semana)):
            ruta_completa = os.path.join(ruta_semana, nombre)
            if os.path.isfile(ruta_completa):
                todos_archivos.append(nombre)
                if nombre.lower().endswith('.xlsx') and not nombre.startswith('~$'):
                    archivos_excel.append(ruta_completa)
    except Exception as e_ls:
        _log_safe('2_Embarques', f'listdir fallo: {e_ls}')

_log_safe('2_Embarques', f'DEBUG Excel | ruta_semana={ruta_semana!r} | todos={todos_archivos} | xlsx={[os.path.basename(p) for p in archivos_excel]}')

if archivos_excel:
    archivo_mas_reciente = max(archivos_excel, key=os.path.getmtime)
    archivo_mas_reciente = archivo_mas_reciente.replace('\\', '/')
    SetVar('v_ruta_excel', archivo_mas_reciente)
    _log_safe('2_Embarques', 'Excel seleccionado: ' + archivo_mas_reciente)
else:
    SetVar('v_ruta_excel', '')
    _log_safe('2_Embarques', 'SIN Excel en ' + ruta_semana + ' | archivos en dir: ' + str(todos_archivos))
