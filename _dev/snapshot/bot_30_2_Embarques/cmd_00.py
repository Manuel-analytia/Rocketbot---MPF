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

# PREFER_PARENT_PATHS
_parent_re = GetVar('v_root_embarques')
if _parent_re and _parent_re != 'ERROR_NOT_VAR' and os.path.isdir(_parent_re):
    root_embarques = _parent_re
else:
    root_embarques = _detectar_root_embarques()
    if not root_embarques:
        _od = os.path.join(os.path.expanduser('~'), 'OneDrive - mpf.com.pe')
        if not os.path.isdir(_od):
            _od = os.path.join(os.path.expanduser('~'), 'OneDrive - mpf.com.pe (1)')
        if os.path.isdir(_od):
            root_embarques = os.path.join(_od, 'SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES')
if root_embarques:
    SetVar('v_root_embarques', root_embarques)
    SetVar('v_ruta_base', root_embarques)
