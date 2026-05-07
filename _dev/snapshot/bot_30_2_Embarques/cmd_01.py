import os

def _detectar_root_embarques(subfolder_default='SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES'):
    username = os.environ.get('USERNAME', '')
    if username.upper() == 'MPFSLY01':
        root_od = r'C:\Users\MPFSLY01\OneDrive - mpf.com.pe (1)'
    else:
        root_od = ''
        for suffix in ['\\OneDrive - mpf.com.pe', '\\OneDrive - mpf.com.pe (1)']:
            candidate = 'C:\\Users\\' + username + suffix
            if os.path.isdir(candidate):
                root_od = candidate
                break
    if not root_od:
        return ''
    # Try Config_Rutas.xlsx for ROOT_EMBARQUES override
    config_path = os.path.join(root_od, subfolder_default, '0. Bot', 'Rutas', 'Config_Rutas.xlsx')
    root_embarques = ''
    if os.path.exists(config_path):
        try:
            import zipfile
            from xml.etree import ElementTree as ET
            ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                  'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                  'p': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            with zipfile.ZipFile(config_path, 'r') as zf:
                shared = []
                if 'xl/sharedStrings.xml' in zf.namelist():
                    for si in ET.fromstring(zf.read('xl/sharedStrings.xml')).findall('m:si', ns):
                        shared.append(''.join(n.text or '' for n in si.findall('.//m:t', ns)))
                wb = ET.fromstring(zf.read('xl/workbook.xml'))
                rels = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
                rmap = {r.get('Id'): r.get('Target') for r in rels.findall('p:Relationship', ns)}
                for sh in wb.findall('m:sheets/m:sheet', ns):
                    if sh.get('name') != 'BASE':
                        continue
                    rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    target = rmap.get(rid, '')
                    if not target:
                        continue
                    sheet_xml = ET.fromstring(zf.read('xl/' + target.lstrip('/')))
                    for row in sheet_xml.findall('m:sheetData/m:row', ns):
                        vals = []
                        for c in row.findall('m:c', ns):
                            t = c.get('t', '')
                            v = c.find('m:v', ns)
                            if v is None:
                                vals.append('')
                            elif t == 's':
                                idx = int(v.text)
                                vals.append(shared[idx] if idx < len(shared) else '')
                            else:
                                vals.append(v.text or '')
                        if len(vals) >= 2 and vals[0] == 'ROOT_EMBARQUES' and vals[1].strip():
                            val = vals[1].strip()
                            root_embarques = val if os.path.isabs(val) else os.path.join(root_od, val)
                            break
        except Exception:
            pass
    if not root_embarques:
        root_embarques = os.path.join(root_od, subfolder_default)
    return root_embarques
def _detectar_semana(root_embarques):
    from datetime import datetime
    semana_actual = datetime.now().isocalendar()[1]
    anio_actual   = datetime.now().isocalendar()[0]
    semana_forzada = ''
    anio_forzado = ''
    if root_embarques:
        pruebas_path = os.path.join(root_embarques, '0. Bot', 'Pruebas', 'pruebas.txt')
        if os.path.exists(pruebas_path):
            try:
                with open(pruebas_path, 'r', encoding='utf-8') as _f:
                    for _ln in _f:
                        _ln = _ln.strip()
                        if not _ln:
                            continue
                        if '|' in _ln:
                            _sw, _sep, _sa = _ln.partition('|')
                            _sw = _sw.strip(); _sa = _sa.strip()
                            if _sw.isdigit() and int(_sw) > 0:
                                semana_forzada = _sw
                                if _sa.isdigit() and int(_sa) >= 2000:
                                    anio_forzado = _sa
                        elif _ln.isdigit() and int(_ln) > 0:
                            semana_forzada = _ln
            except Exception:
                pass
    return (semana_forzada if semana_forzada else str(semana_actual)), (anio_forzado if anio_forzado else str(anio_actual))

from datetime import datetime, date
import os

root_embarques = GetVar('v_root_embarques') or GetVar('v_ruta_base') or ''
numero_semana_trabajo, anio_trabajo = _detectar_semana(root_embarques)
iso_anio = int(anio_trabajo)
nombre_semana = 'Semana ' + numero_semana_trabajo

# Calcular mes del lunes de la semana ISO para nombre de hoja registro
lunes_semana = date.fromisocalendar(iso_anio, int(numero_semana_trabajo), 1)
meses_es = {
    1:'ENERO', 2:'FEBRERO', 3:'MARZO', 4:'ABRIL', 5:'MAYO', 6:'JUNIO',
    7:'JULIO', 8:'AGOSTO', 9:'SETIEMBRE', 10:'OCTUBRE', 11:'NOVIEMBRE', 12:'DICIEMBRE'
}
mes_semana = meses_es[lunes_semana.month]

SetVar('v_anio_semana', str(iso_anio))
SetVar('v_nombre_semana', nombre_semana)
SetVar('v_numero_semana_trabajo', numero_semana_trabajo)
SetVar('v_mes_semana', mes_semana)
SetVar('v_carpeta_mes_actual', 'SEMANAS')
