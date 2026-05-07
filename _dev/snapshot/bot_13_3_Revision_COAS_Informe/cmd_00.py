# AUTODETECT_PROLOGUE: resolver todas las rutas desde expanduser(~)
# asi cualquier PC/usuario corre sin necesidad de editar hardvars.
import os as _pre_os
from datetime import datetime as _pre_dt
_home = _pre_os.path.expanduser("~")
_od = None
_marker_sub = "SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES"
try:
    for _nn in _pre_os.listdir(_home):
        if _nn.lower().startswith("onedrive"):
            _cand = _pre_os.path.join(_home, _nn)
            if _pre_os.path.isdir(_pre_os.path.join(_cand, _marker_sub)):
                _od = _cand
                break
except Exception:
    pass
if not _od:
    # Fallback a los paths conocidos
    _od = _pre_os.path.join(_home, "OneDrive - mpf.com.pe (1)")
    if not _pre_os.path.isdir(_od):
        _od = _pre_os.path.join(_home, "OneDrive - mpf.com.pe")
SetVar("v_root_onedrive", _od)
_re = _pre_os.path.join(_od, "SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES")
SetVar("v_root_embarques", _re)
SetVar("v_ruta_base", _re)
SetVar("v_root_riverside", _pre_os.path.join(_od, "SAN ISIDRO - RIVERSIDE NATURAL FOODS, LTD"))
_br = _pre_os.path.join(_re, "0. Bot")
SetVar("v_bot_root", _br)
SetVar("v_bot_logs", _pre_os.path.join(_br, "Logs"))
SetVar("v_bot_temp", _pre_os.path.join(_br, "Temp"))
SetVar("v_bot_rutas", _pre_os.path.join(_br, "Rutas"))
SetVar("v_bot_correos", _pre_os.path.join(_br, "Correos"))
SetVar("v_bot_pruebas", _pre_os.path.join(_br, "Pruebas"))
SetVar("v_ruta_correos_coas", _pre_os.path.join(_br, "Correos", "COAS_e_Informes"))
SetVar("v_ruta_correos_invoice", _pre_os.path.join(_br, "Correos", "Invoice"))
SetVar("v_ruta_correos_estatus", _pre_os.path.join(_br, "Correos", "Estatus"))
SetVar("v_ruta_config_rutas", _pre_os.path.join(_br, "Rutas", "Config_Rutas.xlsx"))
_anio_pre = str(_pre_dt.now().isocalendar()[0])
_semana_pre = str(_pre_dt.now().isocalendar()[1])
# Override con año y semana forzados de pruebas.txt (linea 'N|YYYY')
try:
    _pruebas_pre = _pre_os.path.join(_br, 'Pruebas', 'pruebas.txt')
    if _pre_os.path.exists(_pruebas_pre):
        with open(_pruebas_pre, 'r', encoding='utf-8') as _pfp:
            for _pl in _pfp:
                _pl = _pl.strip()
                if not _pl:
                    continue
                if '|' in _pl:
                    _sw, _sep, _sa = _pl.partition('|')
                    _sw = _sw.strip(); _sa = _sa.strip()
                    if _sw.isdigit() and int(_sw) > 0:
                        _semana_pre = _sw
                        if _sa.isdigit() and int(_sa) >= 2000:
                            _anio_pre = _sa
                        break
                elif _pl.isdigit() and int(_pl) > 0:
                    _semana_pre = _pl
                    break
except Exception:
    pass
# Propagar forzados a vars para que el codigo principal los lea
SetVar("v_anio_semana", _anio_pre)
SetVar("v_numero_semana_trabajo", _semana_pre)
SetVar("v_nombre_semana", "Semana " + _semana_pre)
SetVar("v_ruta_registro_coordinacion", _pre_os.path.join(_re, "COORDINACION DE EMBARQUES",
                                                          "Registro de COAS e informes completos " + _anio_pre + ".xlsx"))

# Leer credenciales SMTP desde Config_Rutas.xlsx (hoja CREDENCIALES)
_config_path = _pre_os.path.join(_br, "Rutas", "Config_Rutas.xlsx")
if _pre_os.path.isfile(_config_path):
    try:
        from openpyxl import load_workbook as _pre_lw
        _wb_cfg = _pre_lw(_config_path, data_only=True)
        if "CREDENCIALES" in _wb_cfg.sheetnames:
            _ws_cred = _wb_cfg["CREDENCIALES"]
            for _r in _ws_cred.iter_rows(min_row=2, values_only=True):
                if not _r or not _r[0]:
                    continue
                _k = str(_r[0]).strip()
                _v = str(_r[1] or "").strip()
                if _k == "SMTP_HOST" and _v: SetVar("v_smtp_host", _v)
                elif _k == "SMTP_PORT" and _v: SetVar("v_smtp_port", _v)
                elif _k == "SMTP_USER" and _v: SetVar("v_smtp_user", _v)
                elif _k == "SMTP_PASS" and _v: SetVar("v_smtp_pass", _v)
                elif _k == "SMTP_SSL" and _v: SetVar("v_smtp_ssl", _v.lower())
    except Exception:
        pass
# Path para libs locales del bot (pypdf, etc. instalados sin admin)
try:
    import sys as _sys_pre
    _libs_pre = _pre_os.path.join(_br, 'libs')
    if _pre_os.path.isdir(_libs_pre) and _libs_pre not in _sys_pre.path:
        _sys_pre.path.insert(0, _libs_pre)
except Exception:
    pass
# FIN AUTODETECT_PROLOGUE

import os
from datetime import datetime, date

PROCESO = "3_Revision_COAS_Informe"
MESES = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"]

def _log(msg):
    import os as _os
    from datetime import datetime as _dt
    # Resolver log path inline (Rocketbot no ve funciones helper desde otras funciones)
    rp = GetVar("v_orq_log_path") or GetVar("v_ruta_log_actual") or ""
    if not (rp and rp != "ERROR_NOT_VAR" and _os.path.isabs(rp) and _os.path.isdir(_os.path.dirname(rp))):
        rp = ""
        root_emb = GetVar("v_root_embarques") or GetVar("v_ruta_base") or ""
        if root_emb and root_emb != "ERROR_NOT_VAR":
            logs_dir = _os.path.join(root_emb, "0. Bot", "Logs")
            if _os.path.isdir(logs_dir):
                try:
                    logs = [f for f in _os.listdir(logs_dir) if f.startswith("LOG - ") and f.endswith(".txt")]
                    if logs:
                        rp = _os.path.join(logs_dir, sorted(logs)[-1])
                except Exception:
                    pass
    if rp:
        try:
            with open(rp, "a", encoding="utf-8") as f:
                f.write(_dt.now().strftime("%H:%M:%S") + "|" + "3_Revision_COAS_Informe" + "|" + msg + chr(10))
        except Exception:
            pass

def _getv(name, default=""):
    v = GetVar(name)
    if not v or v == "ERROR_NOT_VAR":
        return default
    return v

def _leer_txt(path):
    import os as _os
    items = []
    if path and _os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                for linea in f:
                    s = linea.strip()
                    if s and not s.startswith("#"):
                        items.append(s)
        except Exception:
            pass
    return items

def _buscar_por_embarque(dir_busqueda, embarque):
    import os as _os
    if not dir_busqueda or not _os.path.isdir(dir_busqueda):
        return []
    emb = (embarque or "").strip().lower()
    if not emb:
        return []
    out = []
    try:
        for name in sorted(_os.listdir(dir_busqueda)):
            full = _os.path.join(dir_busqueda, name)
            low = name.lower()
            # Caso A: archivo plano con el embarque en el nombre
            if _os.path.isfile(full):
                if low.endswith(".pdf") and emb in low:
                    out.append(full)
                continue
            # Caso B: subcarpeta cuyo nombre ES exactamente el embarque -> tomar todos los PDFs
            if _os.path.isdir(full) and low.strip() == emb:
                try:
                    for sub in sorted(_os.listdir(full)):
                        sub_full = _os.path.join(full, sub)
                        if _os.path.isfile(sub_full) and sub.lower().endswith(".pdf"):
                            out.append(sub_full)
                except Exception:
                    pass
    except Exception as e:
        try:
            _rp = GetVar("v_orq_log_path") or GetVar("v_ruta_log_actual") or ""
            if _rp and _rp != "ERROR_NOT_VAR" and _os.path.isabs(_rp):
                from datetime import datetime as _bdt
                with open(_rp, "a", encoding="utf-8") as _fh:
                    _fh.write(_bdt.now().strftime("%H:%M:%S") + "|3_Revision_COAS_Informe|WARN listdir fallo en " + dir_busqueda + ": " + str(e) + chr(10))
        except Exception:
            pass
    return out

def _consolidar_pdf(paths_pdfs, excluir_texto, destino):
    """Consolida PDFs truncando desde la pagina que contenga 'excluir_texto'
    (match tolerante a acentos + whitespace). Originales no se modifican."""
    import sys as _sys
    import os as _os
    import unicodedata as _ud

    # Inyectar path del modulo PDF de Rocketbot
    _rb_candidates = []
    try:
        _rb_base = _os.path.dirname(_sys.executable)
        _rb_candidates.append(_os.path.join(_rb_base, "modules", "PDF", "libs", "Windows", "x64" if _sys.maxsize > 2**32 else "x86"))
    except Exception:
        pass
    _rb_candidates.extend([
        r"C:\Program Files (x86)\Rocketbot\modules\PDF\libs\Windows\x64",
        r"C:\Program Files\Rocketbot\modules\PDF\libs\Windows\x64",
    ])
    for _p in _rb_candidates:
        if _os.path.isdir(_p) and _p not in _sys.path:
            _sys.path.insert(0, _p)

    # Normalizar excluir_texto (sin acentos + whitespace colapsado)
    try:
        _en = _ud.normalize("NFKD", excluir_texto or "").encode("ascii", "ignore").decode("ascii").upper()
        excluir_norm = " ".join(_en.split())
    except Exception:
        excluir_norm = " ".join((excluir_texto or "").upper().split())

    def _log_line(msg):
        try:
            _rp = GetVar("v_orq_log_path") or GetVar("v_ruta_log_actual") or ""
            if _rp and _rp != "ERROR_NOT_VAR" and _os.path.isabs(_rp):
                from datetime import datetime as _pdt
                with open(_rp, "a", encoding="utf-8") as _pfh:
                    _pfh.write(_pdt.now().strftime("%H:%M:%S") + "|3_Revision_COAS_Informe|" + msg + chr(10))
        except Exception:
            pass

    # Preferir PyPDF3 (conocida estable, ya en Rocketbot), pypdf, PyPDF2, fitz
    _lib = None
    _Reader = None
    _Writer = None
    for _modname in ("PyPDF3", "pypdf", "PyPDF2"):
        try:
            _mod = __import__(_modname)
            _Reader = _mod.PdfFileReader if _modname == "PyPDF3" else _mod.PdfReader
            _Writer = _mod.PdfFileWriter if _modname == "PyPDF3" else _mod.PdfWriter
            _lib = _modname
            break
        except Exception:
            continue
    if not _lib:
        try:
            import fitz as _fz
            _lib = "fitz"
        except Exception:
            _log_line("ERROR ninguna libreria PDF disponible")
            raise RuntimeError("Sin libreria PDF")

    _log_line("INFO consolidando con lib=" + _lib + " | buscando=" + repr(excluir_norm))

    paginas_omitidas = 0

    if _lib in ("PyPDF3", "pypdf", "PyPDF2"):
        writer = _Writer()
        for ruta in paths_pdfs:
            _nombre = _os.path.basename(ruta)
            try:
                reader = _Reader(ruta)
            except Exception as e:
                _log_line("WARN no se puede abrir " + _nombre + ": " + str(e))
                continue
            try:
                total = len(reader.pages)
            except Exception:
                try:
                    total = reader.getNumPages()
                except Exception:
                    total = 0
            corte = total
            for idx in range(total):
                try:
                    _page = reader.pages[idx] if hasattr(reader, "pages") else reader.getPage(idx)
                    if hasattr(_page, "extract_text"):
                        _txt = _page.extract_text()
                    else:
                        _txt = _page.extractText()
                    _txt = _txt or ""
                    _tn = _ud.normalize("NFKD", _txt).encode("ascii", "ignore").decode("ascii").upper()
                    _txn = " ".join(_tn.split())
                except Exception as _ee:
                    _txn = ""
                    if idx == 0:
                        _log_line("WARN extract_text fallo en p0 de " + _nombre + ": " + str(_ee))
                if excluir_norm in _txn:
                    corte = idx
                    break
            _log_line("INFO " + _nombre + ": total=" + str(total) + " corte=" + str(corte) + " (sin_corte=" + str(corte >= total) + ")")
            for idx in range(corte):
                try:
                    _page = reader.pages[idx] if hasattr(reader, "pages") else reader.getPage(idx)
                    if hasattr(writer, "add_page"):
                        writer.add_page(_page)
                    else:
                        writer.addPage(_page)
                except Exception as e:
                    _log_line("WARN append p" + str(idx) + " de " + _nombre + " fallo: " + str(e))
            paginas_omitidas += (total - corte)
        with open(destino, "wb") as _of:
            writer.write(_of)
        return paginas_omitidas

    # Fallback fitz
    out = _fz.open()
    for ruta in paths_pdfs:
        _nombre = _os.path.basename(ruta)
        try:
            doc = _fz.open(ruta)
        except Exception as e:
            _log_line("WARN no se puede abrir " + _nombre + ": " + str(e))
            continue
        corte = doc.page_count
        for page in doc:
            try:
                _txt_raw = page.get_text() or ""
                _tn = _ud.normalize("NFKD", _txt_raw).encode("ascii", "ignore").decode("ascii").upper()
                _txn = " ".join(_tn.split())
            except Exception as _ee:
                _txn = ""
                if page.number == 0:
                    _log_line("WARN get_text fallo en p0 de " + _nombre + ": " + str(_ee))
            if excluir_norm in _txn:
                corte = page.number
                break
        _log_line("INFO " + _nombre + ": total=" + str(doc.page_count) + " corte=" + str(corte) + " (sin_corte=" + str(corte >= doc.page_count) + ")")
        if corte > 0:
            out.insert_pdf(doc, from_page=0, to_page=corte - 1)
        paginas_omitidas += (doc.page_count - corte)
        doc.close()
    out.save(destino)
    out.close()
    return paginas_omitidas

    # Fallback: PyPDF3 (viene con el modulo PDF) o pypdf/PyPDF2 (si usuario instalo)
    _Reader = None
    _Writer = None
    for _modname in ("PyPDF3", "pypdf", "PyPDF2"):
        try:
            _mod = __import__(_modname)
            _Reader = _mod.PdfFileReader if _modname == "PyPDF3" else _mod.PdfReader
            _Writer = _mod.PdfFileWriter if _modname == "PyPDF3" else _mod.PdfWriter
            _lib = _modname
            break
        except Exception:
            continue

    if not _lib:
        _log_err("ERROR ninguna libreria PDF disponible (fitz/PyPDF3/pypdf/PyPDF2)")
        raise RuntimeError("Sin libreria PDF")

    writer = _Writer()
    for ruta in paths_pdfs:
        try:
            reader = _Reader(ruta)
        except Exception as e:
            _log_err("WARN no se puede abrir PDF " + ruta + ": " + str(e))
            continue
        # Numero de paginas difiere por libreria
        try:
            total = len(reader.pages)
        except Exception:
            try:
                total = reader.getNumPages()
            except Exception:
                total = 0
        corte = total
        if excluir_norm and total:
            for idx in range(total):
                try:
                    _page = reader.pages[idx] if hasattr(reader, "pages") else reader.getPage(idx)
                    _txt = _page.extract_text() if hasattr(_page, "extract_text") else _page.extractText()
                    _txt = _txt or ""
                    _tn = _ud.normalize("NFKD", _txt).encode("ascii", "ignore").decode("ascii").upper()
                    _txn = " ".join(_tn.split())
                except Exception:
                    _txn = ""
                if excluir_norm in _txn:
                    corte = idx
                    _log_err("INFO corte en pag " + str(idx) + " de " + _os.path.basename(ruta))
                    break
        for idx in range(corte):
            try:
                _page = reader.pages[idx] if hasattr(reader, "pages") else reader.getPage(idx)
                if hasattr(writer, "add_page"):
                    writer.add_page(_page)
                else:
                    writer.addPage(_page)
            except Exception as e:
                _log_err("WARN no se puede anexar pagina " + str(idx) + " de " + ruta + ": " + str(e))
        paginas_omitidas += (total - corte)
    with open(destino, "wb") as _out_f:
        writer.write(_out_f)
    return paginas_omitidas

    # Fallback fitz (por si se instala luego)
    out = _fz.open()
    for ruta in paths_pdfs:
        try:
            doc = _fz.open(ruta)
        except Exception as e:
            _log_err("WARN no se puede abrir PDF " + ruta + ": " + str(e))
            continue
        corte = doc.page_count
        if excluir_norm:
            for page in doc:
                try:
                    _txt_raw = page.get_text() or ""
                    texto_n = _ud.normalize("NFKD", _txt_raw).encode("ascii", "ignore").decode("ascii").upper()
                except Exception:
                    texto_n = ""
                if excluir_norm in texto_n:
                    corte = page.number
                    break
        if corte > 0:
            out.insert_pdf(doc, from_page=0, to_page=corte - 1)
        paginas_omitidas += (doc.page_count - corte)
        doc.close()
    out.save(destino)
    out.close()
    return paginas_omitidas

def _enviar_correo(asunto, cuerpo, para, cc_list, co_list, adjunto):
    """Retorna (ok: bool, err: str). En exito err='', en fallo err=descripcion."""
    import os as _os
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    if not para:
        return (False, "sin destinatarios en Para.txt")
    host = (GetVar("v_smtp_host") or "smtp.office365.com") if (GetVar("v_smtp_host") not in ("", "ERROR_NOT_VAR")) else "smtp.office365.com"
    _p = GetVar("v_smtp_port"); port = int(_p) if (_p and _p != "ERROR_NOT_VAR") else 587
    user = GetVar("v_smtp_user") if (GetVar("v_smtp_user") not in ("", "ERROR_NOT_VAR")) else ""
    pwd  = GetVar("v_smtp_pass") if (GetVar("v_smtp_pass") not in ("", "ERROR_NOT_VAR")) else ""
    _s = GetVar("v_smtp_ssl"); ssl_mode = (_s.lower() == "true") if (_s and _s != "ERROR_NOT_VAR") else False
    if not user:
        return (False, "v_smtp_user vacio")
    if not pwd:
        return (False, "v_smtp_pass vacio")
    _stage = "build_msg"
    try:
        msg = MIMEMultipart()
        msg["From"] = user
        msg["To"] = ", ".join(para)
        if cc_list:
            msg["Cc"] = ", ".join(cc_list)
        msg["Subject"] = asunto
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))
        if adjunto and _os.path.exists(adjunto):
            with open(adjunto, "rb") as fh:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(fh.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            "attachment; filename=\"" + _os.path.basename(adjunto) + "\"")
            msg.attach(part)
        destinos = para + cc_list + co_list
        _stage = "connect " + host + ":" + str(port) + (" SSL" if ssl_mode else " STARTTLS")
        if ssl_mode:
            smtp = smtplib.SMTP_SSL(host, port, timeout=300)
        else:
            smtp = smtplib.SMTP(host, port, timeout=300)
            _stage = "starttls"
            smtp.starttls()
        _stage = "login " + user
        smtp.login(user, pwd)
        _att_mb = ""
        if adjunto and _os.path.exists(adjunto):
            try:
                _att_mb = " (adjunto " + str(round(_os.path.getsize(adjunto)/(1024*1024), 2)) + "MB)"
            except Exception:
                pass
        _stage = "sendmail to " + str(len(destinos)) + " dest" + _att_mb
        smtp.sendmail(user, destinos, msg.as_string())
        smtp.quit()
        return (True, "")
    except Exception as e:
        return (False, _stage + " -> " + type(e).__name__ + ": " + str(e))

# ============ FLUJO PRINCIPAL ============

_log("Inicio bot 3_Revision_COAS_Informe")

try:
    from openpyxl import load_workbook
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

registro = _getv("v_ruta_registro_coordinacion")
if not _OPENPYXL_OK:
    _log("ERROR openpyxl no disponible")
    SetVar("v_h3_estado", "ERROR")
    SetVar("v_h3_detalle", "openpyxl no disponible")
elif not registro or not os.path.exists(registro):
    _log("ERROR Registro no existe: " + str(registro))
    SetVar("v_h3_estado", "ERROR")
    SetVar("v_h3_detalle", "No existe Registro " + str(registro))
else:
    _log("Registro: " + registro)
    root_od = _getv("v_root_onedrive")
    coas_fuente = _getv("v_root_coas_fuente", "Archivos de Juliana Legua - ORGANICO")
    inf_prefijo = _getv("v_root_informes_prefijo", "PISCO - ")
    anio = _getv("v_anio_semana", str(datetime.now().isocalendar()[0]))
    pag_excluir = _getv("v_pagina_excluir_informes", "TERMINOS Y CONDICIONES DE LOS SERVICIOS")
    dir_coas = os.path.join(root_od, coas_fuente) if root_od else coas_fuente
    dir_inf  = os.path.join(root_od, inf_prefijo + anio) if root_od else (inf_prefijo + anio)
    _log("COAS dir: " + dir_coas + " | exists=" + str(os.path.isdir(dir_coas)))
    _log("Informes dir: " + dir_inf + " | exists=" + str(os.path.isdir(dir_inf)))
    temp_dir = _getv("v_bot_temp")
    if not temp_dir:
        temp_dir = os.path.join(_getv("v_root_embarques"), "0. Bot", "Temp")
    try:
        os.makedirs(temp_dir, exist_ok=True)
    except Exception:
        pass

    correos_dir = _getv("v_ruta_correos_coas")
    para = _leer_txt(os.path.join(correos_dir, "Para.txt")) if correos_dir else []
    cc_list = _leer_txt(os.path.join(correos_dir, "CC.txt")) if correos_dir else []
    co_list = _leer_txt(os.path.join(correos_dir, "CO.txt")) if correos_dir else []
    # Destinatarios de alertas (Estatus) - para correo de faltantes
    estatus_dir = _getv("v_ruta_correos_estatus")
    para_est = _leer_txt(os.path.join(estatus_dir, "Para.txt")) if estatus_dir else []
    cc_est   = _leer_txt(os.path.join(estatus_dir, "CC.txt"))   if estatus_dir else []
    co_est   = _leer_txt(os.path.join(estatus_dir, "CO.txt"))   if estatus_dir else []
    _log("Correos: Para=" + str(len(para)) + " CC=" + str(len(cc_list)) + " CO=" + str(len(co_list)))

    wb = load_workbook(registro)
    total_filas = 0
    sin_completo = 0
    ya_enviado = 0
    procesados = 0
    enviados = 0
    faltantes = 0
    errores = 0
    alerta_no_completo = []  # filas sin Completo con mas de 4 dias
    for hoja_mes in MESES:
        if hoja_mes not in wb.sheetnames:
            continue
        ws = wb[hoja_mes]
        max_row = ws.max_row
        for row_idx in range(2, max_row + 1):
            cliente  = str(ws.cell(row=row_idx, column=2).value or "").strip()
            embarque = str(ws.cell(row=row_idx, column=3).value or "").strip()
            fecha_raw = ws.cell(row=row_idx, column=4).value
            pedido   = str(ws.cell(row=row_idx, column=5).value or "").strip()
            coas     = str(ws.cell(row=row_idx, column=6).value or "").strip().upper()
            informes = str(ws.cell(row=row_idx, column=7).value or "").strip().upper()
            fecha_coas = str(ws.cell(row=row_idx, column=11).value or "").strip()
            if not embarque:
                continue
            total_filas += 1
            if fecha_coas:
                ya_enviado += 1
                continue
            # Parsear fecha embarque para calcular dias de atraso
            _fec_emb = None
            if fecha_raw:
                if hasattr(fecha_raw, "strftime"):
                    _fec_emb = fecha_raw.date() if hasattr(fecha_raw, "date") else fecha_raw
                else:
                    for _fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                        try:
                            from datetime import datetime as _dtx
                            _fec_emb = _dtx.strptime(str(fecha_raw).strip(), _fmt).date()
                            break
                        except Exception:
                            _fec_emb = None
            _dias_atraso = 0
            if _fec_emb:
                _dias_atraso = (date.today() - _fec_emb).days
            if coas != "COMPLETO" or informes != "COMPLETO":
                sin_completo += 1
                # Alerta: no completos con mas de 4 dias de atraso
                if _dias_atraso > 4:
                    alerta_no_completo.append((cliente, embarque, str(_fec_emb) if _fec_emb else str(fecha_raw), pedido, _dias_atraso, coas, informes))
                continue
            procesados += 1

            archivos_coas = _buscar_por_embarque(dir_coas, embarque)
            archivos_inf  = _buscar_por_embarque(dir_inf, embarque)
            if archivos_coas and archivos_inf:
                existencia = "OK"
            elif not archivos_coas and not archivos_inf:
                existencia = "AMBOS_NO_ENCONTRADOS"
            elif not archivos_coas:
                existencia = "COAS_NO_ENCONTRADO"
            else:
                existencia = "INFORMES_NO_ENCONTRADO"
            ws.cell(row=row_idx, column=10, value=existencia)
            _log(hoja_mes + "/" + embarque + " | COMPLETO | coas=" + str(len(archivos_coas)) + " inf=" + str(len(archivos_inf)) + " -> " + existencia)

            if existencia != "OK":
                faltantes += 1
                est_coas = "NO_ENCONTRADO" if existencia in ("COAS_NO_ENCONTRADO", "AMBOS_NO_ENCONTRADOS") else "OK"
                est_inf  = "NO_ENCONTRADO" if existencia in ("INFORMES_NO_ENCONTRADO", "AMBOS_NO_ENCONTRADOS") else "OK"
                # Normalizar pedido (quitar prefijo 'PO-' o 'PO ')
                _ped_num = pedido.strip()
                _pu = _ped_num.upper()
                if _pu.startswith("PO-"):
                    _ped_num = _ped_num[3:].strip()
                elif _pu.startswith("PO "):
                    _ped_num = _ped_num[3:].strip()
                asunto = "COMEX | Documentos no encontrados - Embarque " + embarque
                _cli_clean_f = cliente.rstrip(". ").strip()
                cuerpo = ("Estimados," + chr(10) + chr(10)
                         + "El embarque " + embarque + " (Cliente " + _cli_clean_f + ", PO " + _ped_num + ") fue marcado" + chr(10)
                         + "como COMPLETO en el Registro de Coordinacion, pero no se encontraron" + chr(10)
                         + "los archivos correspondientes en las rutas configuradas." + chr(10) + chr(10)
                         + "Resultado de la busqueda:" + chr(10)
                         + "  - COAs     : " + est_coas + chr(10)
                         + "    Ruta     : " + dir_coas + chr(10)
                         + "    Criterio : archivos PDF cuyo nombre contenga " + embarque + chr(10)
                         + "  - Informes : " + est_inf + chr(10)
                         + "    Ruta     : " + dir_inf + chr(10)
                         + "    Criterio : subcarpeta con nombre " + embarque + chr(10) + chr(10)
                         + "Favor verificar la existencia y nomenclatura de los archivos/carpetas." + chr(10) + chr(10)
                         + "Saludos," + chr(10) + "Bot COMEX MPF" + chr(10))
                _ok_e, _err_e = _enviar_correo(asunto, cuerpo, para_est, cc_est, co_est, None)
                if _ok_e:
                    _log(hoja_mes + "/" + embarque + " | correo faltantes enviado")
                else:
                    _log(hoja_mes + "/" + embarque + " | WARN correo faltantes fallo: " + _err_e)
                continue

            pdf_dest = os.path.join(temp_dir, embarque + " - Quality Certificate & Test Reports.pdf")
            # Borrar copia previa en temp para evitar mezcla con corrida anterior.
            # Reintentos por locks de sync de OneDrive.
            if os.path.exists(pdf_dest):
                _del_ok = False
                _del_err = ""
                for _intento_del in range(3):
                    try:
                        os.remove(pdf_dest)
                        _del_ok = True
                        break
                    except Exception as _ec_del:
                        _del_err = type(_ec_del).__name__ + ": " + str(_ec_del)
                        if _intento_del < 2:
                            try:
                                import time as _tm_del
                                _tm_del.sleep(1.0)
                            except Exception:
                                pass
                if _del_ok:
                    _log(hoja_mes + "/" + embarque + " | PDF previo eliminado de Temp: " + pdf_dest)
                else:
                    _log(hoja_mes + "/" + embarque + " | WARN no se pudo borrar PDF previo: " + _del_err + " -> se intentara sobrescribir")
            try:
                _consolidar_pdf(archivos_coas + archivos_inf, pag_excluir, pdf_dest)
                _log(hoja_mes + "/" + embarque + " | PDF consolidado en " + pdf_dest)
            except Exception as e:
                _log("ERROR consolidando PDF " + embarque + ": " + str(e))
                errores += 1
                continue

            # Copia consolidado a Riverside SIEMPRE (independiente del envio):
            # si el envio del correo falla, igual queremos dejar el PDF en su carpeta.
            _root_rs = _getv("v_root_riverside")
            if not _root_rs:
                _log(hoja_mes + "/" + embarque + " | WARN v_root_riverside no definido, no se copia consolidado")
            else:
                # Buscar carpeta del embarque por prefijo. Patron real:
                # '<emb> RIVERSIDE <MMM> <YY>'  (ej. '2000009342 RIVERSIDE MAR 26')
                _carpeta_emb_rs = None
                if os.path.isdir(_root_rs):
                    # Match exacto primero
                    _p_exact = os.path.join(_root_rs, embarque)
                    if os.path.isdir(_p_exact):
                        _carpeta_emb_rs = _p_exact
                    else:
                        try:
                            import re as _re_rs
                            for _name_rs in os.listdir(_root_rs):
                                _full_rs = os.path.join(_root_rs, _name_rs)
                                if not os.path.isdir(_full_rs):
                                    continue
                                if _re_rs.match(r"^" + _re_rs.escape(embarque) + r"(\s|\+|$)", _name_rs):
                                    _carpeta_emb_rs = _full_rs
                                    break
                                _toks_rs = _re_rs.split(r"[\s+]+", _name_rs)
                                if embarque in _toks_rs:
                                    _carpeta_emb_rs = _full_rs
                                    break
                        except Exception:
                            pass
                # Si no se encontro carpeta existente, construir nombre esperado
                # con patron '<emb> RIVERSIDE <MMM> <YY>' y crearla.
                if not _carpeta_emb_rs and os.path.isdir(_root_rs):
                    _MESES_RS = {1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",
                                 7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"}
                    if _fec_emb:
                        _car_name_rs = (embarque + " RIVERSIDE " + _MESES_RS[_fec_emb.month]
                                        + " " + str(_fec_emb.year)[-2:])
                    else:
                        _car_name_rs = embarque + " RIVERSIDE"
                    _carpeta_emb_rs = os.path.join(_root_rs, _car_name_rs)
                    try:
                        os.makedirs(_carpeta_emb_rs, exist_ok=True)
                        _log(hoja_mes + "/" + embarque + " | Carpeta Riverside creada: " + _carpeta_emb_rs)
                    except Exception as _ec_mk_rs:
                        _log(hoja_mes + "/" + embarque + " | WARN no se pudo crear carpeta Riverside: " + str(_ec_mk_rs))
                        _carpeta_emb_rs = None
                if not _carpeta_emb_rs:
                    _log(hoja_mes + "/" + embarque + " | WARN carpeta embarque no resuelta en Riverside (root no existe?)")
                elif (not os.path.isfile(pdf_dest)) or os.path.getsize(pdf_dest) == 0:
                    _log(hoja_mes + "/" + embarque + " | WARN consolidado origen invalido o vacio: " + pdf_dest)
                else:
                    _dest_rs = os.path.join(_carpeta_emb_rs, embarque + " - Quality Certificate & Test Reports.pdf")
                    _dest_tmp = _dest_rs + ".part"
                    _ok_copia = False
                    _err_copia = ""
                    _src_size = os.path.getsize(pdf_dest)
                    for _intento in range(3):
                        try:
                            import shutil as _sh
                            if os.path.exists(_dest_tmp):
                                try: os.remove(_dest_tmp)
                                except Exception: pass
                            _sh.copy2(pdf_dest, _dest_tmp)
                            if os.path.getsize(_dest_tmp) != _src_size:
                                raise IOError("tamanio destino " + str(os.path.getsize(_dest_tmp)) + " != origen " + str(_src_size))
                            os.replace(_dest_tmp, _dest_rs)
                            _ok_copia = True
                            break
                        except Exception as _ec:
                            _err_copia = type(_ec).__name__ + ": " + str(_ec)
                            try:
                                if os.path.exists(_dest_tmp): os.remove(_dest_tmp)
                            except Exception:
                                pass
                            if _intento < 2:
                                try:
                                    import time as _tm
                                    _tm.sleep(1.5)
                                except Exception:
                                    pass
                    if _ok_copia:
                        _log(hoja_mes + "/" + embarque + " | Consolidado copiado a Riverside: " + _dest_rs)
                    else:
                        _log(hoja_mes + "/" + embarque + " | WARN copia Riverside fallo tras 3 intentos: " + _err_copia)

            # Normalizar pedido (quitar prefijo 'PO-' o 'PO ') para no duplicar
            _ped_num_e = pedido.strip()
            _pu_e = _ped_num_e.upper()
            if _pu_e.startswith("PO-"):
                _ped_num_e = _ped_num_e[3:].strip()
            elif _pu_e.startswith("PO "):
                _ped_num_e = _ped_num_e[3:].strip()
            asunto_env = "PO-" + _ped_num_e + " || OP-" + embarque + " || Riverside"
            _cli_clean_e = cliente.rstrip(". ").strip()
            # Plantilla en INGLES (envio del consolidado COAs + Informes)
            cuerpo_env = ("Dear all," + chr(10) + chr(10)
                         + "Please find attached the consolidated Quality Certificates and Test Reports" + chr(10)
                         + "document for shipment " + embarque + " from " + _cli_clean_e + " (PO " + _ped_num_e + ")." + chr(10) + chr(10)
                         + "Best regards," + chr(10) + "COMEX MPF Bot" + chr(10))
            _ok_env, _err_env = _enviar_correo(asunto_env, cuerpo_env, para, cc_list, co_list, pdf_dest)
            if _ok_env:
                from datetime import timezone as _tz, timedelta as _td
                _ts_lima = datetime.now(_tz(_td(hours=-5))).strftime("%d/%m/%Y %H:%M:%S")
                ws.cell(row=row_idx, column=11, value=_ts_lima)
                enviados += 1
                _log(hoja_mes + "/" + embarque + " | COAs enviado")
            else:
                errores += 1
                _log(hoja_mes + "/" + embarque + " | Falla envio correo: " + _err_env)
    # Alerta HTML a Estatus: no completos con >4 dias
    if alerta_no_completo and para_est:
        _rows_a = []
        for (cli_a, emb_a, fec_a, ped_a, dias_a, coas_a, inf_a) in alerta_no_completo:
            _cli_ah = cli_a.replace("&","&amp;").replace("<","&lt;")
            _color_a = "#ffe0e0" if dias_a > 8 else "#fff8e0"
            _rows_a.append(
                "<tr style=\"background:" + _color_a + ";\">"
                "<td style=\"padding:6px 10px;font-family:monospace;\">" + emb_a + "</td>"
                "<td style=\"padding:6px 10px;\">" + _cli_ah + "</td>"
                "<td style=\"padding:6px 10px;\">" + fec_a + "</td>"
                "<td style=\"padding:6px 10px;text-align:center;\"><b>" + str(dias_a) + "</b></td>"
                "<td style=\"padding:6px 10px;font-family:monospace;\">" + ped_a + "</td>"
                "<td style=\"padding:6px 10px;text-align:center;\">" + (coas_a or "(vacio)") + "</td>"
                "<td style=\"padding:6px 10px;text-align:center;\">" + (inf_a or "(vacio)") + "</td>"
                "</tr>"
            )
        _html_a = (
            "<html><body style=\"font-family:Calibri,Arial,sans-serif;font-size:13px;color:#222;\">"
            "<p>Estimados,</p>"
            "<p>Se detectaron <b>" + str(len(alerta_no_completo)) + " embarques</b> con mas de 4 dias de atraso "
            "que aun no estan marcados como <b>Completo</b> en COAs e Informes.</p>"
            "<table cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;border:1px solid #999;\">"
            "<thead><tr style=\"background:#333;color:#fff;\">"
            "<th style=\"padding:6px 10px;\">Embarque</th>"
            "<th style=\"padding:6px 10px;\">Cliente</th>"
            "<th style=\"padding:6px 10px;\">Fecha</th>"
            "<th style=\"padding:6px 10px;\">Dias</th>"
            "<th style=\"padding:6px 10px;\">PO</th>"
            "<th style=\"padding:6px 10px;\">COAs</th>"
            "<th style=\"padding:6px 10px;\">Informes</th>"
            "</tr></thead><tbody>"
            + "".join(_rows_a) +
            "</tbody></table>"
            "<p>Por favor marcar las columnas COAs e Informes como <b>Completo</b> cuando la documentacion este lista.</p>"
            "<p>Saludos,<br><i>Bot COMEX MPF</i></p>"
            "</body></html>"
        )
        _asunto_a = "COMEX | Embarques sin Completo (>4 dias) - " + str(date.today())
        # SMTP inline
        import smtplib as _smtp_a
        from email.mime.text import MIMEText as _MT_a
        from email.mime.multipart import MIMEMultipart as _MM_a
        _host_a = GetVar("v_smtp_host") or "smtp.office365.com"
        _port_a = int(GetVar("v_smtp_port") or "587") if (GetVar("v_smtp_port") not in ("", "ERROR_NOT_VAR")) else 587
        _user_a = GetVar("v_smtp_user") if GetVar("v_smtp_user") not in ("", "ERROR_NOT_VAR") else ""
        _pwd_a = GetVar("v_smtp_pass") if GetVar("v_smtp_pass") not in ("", "ERROR_NOT_VAR") else ""
        _ssl_a_raw = GetVar("v_smtp_ssl") or "false"
        _ssl_a = (_ssl_a_raw.lower() == "true") if (_ssl_a_raw and _ssl_a_raw != "ERROR_NOT_VAR") else False
        try:
            _m_a = _MM_a("alternative")
            _m_a["From"] = _user_a
            _m_a["To"] = ", ".join(para_est)
            if cc_est:
                _m_a["Cc"] = ", ".join(cc_est)
            _m_a["Subject"] = _asunto_a
            _m_a.attach(_MT_a(_html_a, "html", "utf-8"))
            _dest_a = para_est + cc_est + co_est
            if _ssl_a:
                _sm_a = _smtp_a.SMTP_SSL(_host_a, _port_a, timeout=30)
            else:
                _sm_a = _smtp_a.SMTP(_host_a, _port_a, timeout=30)
                _sm_a.starttls()
            if _user_a and _pwd_a:
                _sm_a.login(_user_a, _pwd_a)
            _sm_a.sendmail(_user_a, _dest_a, _m_a.as_string())
            _sm_a.quit()
            _log("Alerta no-completos enviada a " + str(_dest_a) + " (" + str(len(alerta_no_completo)) + " embarques)")
        except Exception as _ea:
            _log("ERROR enviando alerta no-completos: " + type(_ea).__name__ + ": " + str(_ea))

    try:
        wb.save(registro)
        _log("Registro guardado")
    except Exception as e:
        _log("ERROR guardando Registro: " + str(e))
        errores += 1

    detalle = ("total=" + str(total_filas)
              + " ya_enviado=" + str(ya_enviado)
              + " sin_completo=" + str(sin_completo)
              + " procesados=" + str(procesados)
              + " enviados=" + str(enviados)
              + " faltantes=" + str(faltantes)
              + " errores=" + str(errores))
    SetVar("v_h3_estado", "OK" if errores == 0 else "OK_CON_ERRORES")
    SetVar("v_h3_detalle", detalle)
    _log("Resumen | " + detalle)

_log("Fin bot 3_Revision_COAS_Informe")
