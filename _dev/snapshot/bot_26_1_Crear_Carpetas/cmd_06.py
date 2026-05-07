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

# Leer todos los valores intermedios para debug
ruta_base   = GetVar('v_ruta_base') or ''
anio_semana = GetVar('v_anio_semana') or ''
ruta_anio   = GetVar('v_ruta_anio') or ''
nombre_sem  = GetVar('v_nombre_semana') or ''
ruta_semana = GetVar('v_ruta_semana') or ''

_log_safe('1_Crear_Carpetas', f'DEBUG rutas | base={ruta_base!r} | anio={anio_semana!r} | ruta_anio={ruta_anio!r} | nombre_sem={nombre_sem!r} | ruta_semana={ruta_semana!r}')

debug_vars = f"base={ruta_base!r} | anio={anio_semana!r} | ruta_anio={ruta_anio!r} | nombre_sem={nombre_sem!r} | ruta_semana={ruta_semana!r}"

# Validar que la ruta sea absoluta antes de usarla
if not ruta_semana or not os.path.isabs(ruta_semana):
    SetVar('v_ruta_excel', '')
    SetVar('v_h1_estado', 'OK_SIN_EXCEL')
    SetVar('v_h1_detalle', 'RUTA_SEMANA_INVALIDA | ' + debug_vars)
else:
    try:
        os.makedirs(ruta_semana, exist_ok=True)
    except Exception as e_mk:
        _log_safe('1_Crear_Carpetas', f'makedirs fallo: {e_mk}')

    todos_archivos = []
    archivos_excel = []
    if os.path.isdir(ruta_semana):
        try:
            for nombre in sorted(os.listdir(ruta_semana)):
                ruta_completa = os.path.join(ruta_semana, nombre)
                if os.path.isfile(ruta_completa):
                    todos_archivos.append(nombre)
                    if nombre.lower().endswith('.xlsx') and not nombre.startswith('~$'):
                        archivos_excel.append(ruta_completa)
        except Exception as e_ls:
            _log_safe('1_Crear_Carpetas', f'listdir fallo: {e_ls}')

    _log_safe('1_Crear_Carpetas', f'DEBUG Excel search | ruta_semana={ruta_semana!r} | existe={os.path.isdir(ruta_semana)} | todos={todos_archivos} | xlsx={[os.path.basename(p) for p in archivos_excel]}')

    if archivos_excel:
        archivo_mas_reciente = max(archivos_excel, key=os.path.getmtime)
        SetVar('v_ruta_excel', archivo_mas_reciente)
        SetVar('v_h1_estado', 'OK')
        SetVar('v_h1_detalle', 'Excel detectado: ' + os.path.basename(archivo_mas_reciente))
        _log_safe('1_Crear_Carpetas', 'Excel seleccionado: ' + archivo_mas_reciente)
    else:
        SetVar('v_ruta_excel', '')
        SetVar('v_h1_estado', 'OK_SIN_EXCEL')
        SetVar('v_h1_detalle', 'SIN_EXCEL en ' + ruta_semana + ' | archivos=' + str(todos_archivos) + ' | ' + debug_vars)
