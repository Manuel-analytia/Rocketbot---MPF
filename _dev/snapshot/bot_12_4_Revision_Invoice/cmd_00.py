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
import re
from datetime import datetime, date

PROCESO = "4_Revision_Invoice"
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
                f.write(_dt.now().strftime("%H:%M:%S") + "|" + "4_Revision_Invoice" + "|" + msg + chr(10))
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

def _parse_fecha(valor):
    from datetime import datetime as _dt, date as _date
    if not valor:
        return None
    if isinstance(valor, _dt):
        return valor.date()
    if isinstance(valor, _date):
        return valor
    s = str(valor).strip()
    for fmt in ("%d/%m/%Y","%d-%m-%Y","%Y-%m-%d","%d/%m/%y","%d-%m-%y"):
        try:
            return _dt.strptime(s, fmt).date()
        except Exception:
            continue
    return None

def _detectar_po(nombre_archivo):
    import re as _re
    low = nombre_archivo.lower()
    m = _re.search(r"invoice\s+po\s+(\S+)", low)
    if m:
        po_num = m.group(1).replace(".pdf", "").strip()
        return True, po_num
    return False, ""

def _buscar_invoices(carpeta_embarque, embarque):
    import os as _os
    import re as _re
    if not _os.path.isdir(carpeta_embarque):
        return []
    emb_low = embarque.lower().strip()
    found = []
    try:
        for name in sorted(_os.listdir(carpeta_embarque)):
            full = _os.path.join(carpeta_embarque, name)
            if not _os.path.isfile(full):
                continue
            low = name.lower()
            if not low.endswith(".pdf"):
                continue
            if emb_low not in low:
                continue
            if "invoice" not in low:
                continue
            if "inbox" in low:
                continue
            # Inline detectar PO (no podemos llamar _detectar_po desde funcion - REGLA 6)
            _mpo = _re.search(r"invoice\s+po\s+(\S+)", low)
            if _mpo:
                es_po = True
                po_num = _mpo.group(1).replace(".pdf", "").strip()
            else:
                es_po = False
                po_num = ""
            found.append((full, name, es_po, po_num))
    except Exception as e:
        # Inline log write (REGLA 6: no podemos llamar _log desde funciones)
        try:
            _rp = GetVar("v_orq_log_path") or GetVar("v_ruta_log_actual") or ""
            if _rp and _rp != "ERROR_NOT_VAR" and _os.path.isabs(_rp):
                from datetime import datetime as _bdt
                with open(_rp, "a", encoding="utf-8") as _fh:
                    _fh.write(_bdt.now().strftime("%H:%M:%S") + "|4_Revision_Invoice|WARN listdir fallo en " + carpeta_embarque + ": " + str(e) + chr(10))
        except Exception:
            pass
    return found


def _encontrar_carpeta_embarque(dir_base, embarque):
    """Busca la carpeta cuyo nombre empiece con el numero de embarque.
    Maneja casos: '<emb>', '<emb> RIVERSIDE ...', '<emb>+<otros> RIVERSIDE...',
    '<emb>  RIVERSIDE...' (doble espacio)."""
    import os as _os
    import re as _re
    emb = embarque.strip()
    if not _os.path.isdir(dir_base):
        return None
    # Match exacto primero
    p = _os.path.join(dir_base, emb)
    if _os.path.isdir(p):
        return p
    try:
        for name in _os.listdir(dir_base):
            full = _os.path.join(dir_base, name)
            if not _os.path.isdir(full):
                continue
            # Token del embarque al inicio seguido de espacio, + o fin
            if _re.match(r"^" + _re.escape(emb) + r"(\s|\+|$)", name):
                return full
            # Caso compuesto: "<emb1> + <emb2> ..." donde emb aparece como token
            toks = _re.split(r"[\s+]+", name)
            if emb in toks:
                return full
    except Exception:
        pass
    return None

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

_log("Inicio bot 4_Revision_Invoice")

try:
    from openpyxl import load_workbook
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

registro = _getv("v_ruta_registro_coordinacion")
if not _OPENPYXL_OK:
    _log("ERROR openpyxl no disponible")
    SetVar("v_h4_estado", "ERROR")
    SetVar("v_h4_detalle", "openpyxl no disponible")
elif not registro or not os.path.exists(registro):
    _log("ERROR Registro no existe: " + str(registro))
    SetVar("v_h4_estado", "ERROR")
    SetVar("v_h4_detalle", "No existe Registro " + str(registro))
else:
    _log("Registro: " + registro)
    root_od = _getv("v_root_onedrive")
    root_rs_rel = _getv("v_root_riverside", "SAN ISIDRO - RIVERSIDE NATURAL FOODS, LTD")
    if os.path.isabs(root_rs_rel):
        dir_riverside = root_rs_rel
    else:
        dir_riverside = os.path.join(root_od, root_rs_rel) if root_od else root_rs_rel
    _log("Riverside dir: " + dir_riverside + " | exists=" + str(os.path.isdir(dir_riverside)))

    correos_dir = _getv("v_ruta_correos_invoice")
    para = _leer_txt(os.path.join(correos_dir, "Para.txt")) if correos_dir else []
    cc_list = _leer_txt(os.path.join(correos_dir, "CC.txt")) if correos_dir else []
    co_list = _leer_txt(os.path.join(correos_dir, "CO.txt")) if correos_dir else []
    # Destinatarios de alertas (Estatus) - para correo de pendientes
    estatus_dir = _getv("v_ruta_correos_estatus")
    para_est = _leer_txt(os.path.join(estatus_dir, "Para.txt")) if estatus_dir else []
    cc_est   = _leer_txt(os.path.join(estatus_dir, "CC.txt"))   if estatus_dir else []
    co_est   = _leer_txt(os.path.join(estatus_dir, "CO.txt"))   if estatus_dir else []
    _log("Correos Invoice: Para=" + str(len(para)) + " CC=" + str(len(cc_list)) + " CO=" + str(len(co_list)))

    hoy = date.today()
    _log("Hoy: " + str(hoy))

    wb = load_workbook(registro)
    total_filas = 0
    ya_enviado = 0
    futuro = 0
    mismo_dia = 0
    procesados = 0
    enviados = 0
    sin_archivo = 0
    atrasados = 0
    errores = 0
    pendientes = []  # lista de (cliente, embarque, fecha, pedido, dias, carpeta)

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
            fecha_envio_inv = str(ws.cell(row=row_idx, column=9).value or "").strip()
            if not embarque:
                continue
            total_filas += 1
            if fecha_envio_inv:
                ya_enviado += 1
                continue
            fecha_emb = _parse_fecha(fecha_raw)
            if not fecha_emb:
                _log(hoja_mes + "/" + embarque + " | fecha invalida: " + repr(fecha_raw))
                continue
            if fecha_emb > hoy:
                futuro += 1
                continue
            dias = (hoy - fecha_emb).days
            if dias < 1:
                mismo_dia += 1
                continue
            procesados += 1
            retraso = dias > 1
            if retraso:
                _log("WARNING INVOICE_ATRASADO embarque=" + embarque + " fecha=" + str(fecha_emb) + " dias_retraso=" + str(dias))
                atrasados += 1

            carpeta_emb_real = _encontrar_carpeta_embarque(dir_riverside, embarque)
            carpeta_emb = carpeta_emb_real if carpeta_emb_real else os.path.join(dir_riverside, embarque)
            invoices = _buscar_invoices(carpeta_emb, embarque) if carpeta_emb_real else []
            _log(hoja_mes + "/" + embarque + " | fecha=" + str(fecha_emb) + " dias=" + str(dias) + " | invoices=" + str(len(invoices)) + " | carpeta=" + (os.path.basename(carpeta_emb_real) if carpeta_emb_real else "(no encontrada)"))

            if not invoices:
                sin_archivo += 1
                # Si no se encontro carpeta real, pasar "" para que el display lo marque
                pendientes.append((cliente, embarque, str(fecha_emb), pedido, dias, carpeta_emb_real if carpeta_emb_real else ""))
                continue

            enviados_fila = 0
            # Normalizar pedido (quitar prefijo 'PO-' o 'PO ') para no duplicar en asunto
            _ped_num = pedido.strip()
            _pu = _ped_num.upper()
            if _pu.startswith("PO-"):
                _ped_num = _ped_num[3:].strip()
            elif _pu.startswith("PO "):
                _ped_num = _ped_num[3:].strip()
            for (path, nombre, es_po, po_num) in invoices:
                if es_po:
                    _po_subj = (po_num or _ped_num)
                    asunto = "PO-" + _po_subj + " / OP-" + embarque
                    bloque_po = " (PO " + _ped_num + ")"
                else:
                    asunto = "PO-" + _ped_num + " / OP-" + embarque
                    bloque_po = ""
                _cli_clean = cliente.rstrip(". ").strip()
                # Plantilla en INGLES (envio del Invoice)
                cuerpo = ("Dear all," + chr(10) + chr(10)
                         + "Please find attached the Invoice for shipment " + embarque + " from " + _cli_clean + bloque_po + "." + chr(10) + chr(10)
                         + "Best regards," + chr(10) + "COMEX MPF Bot" + chr(10))
                ok, _err_inv = _enviar_correo(asunto, cuerpo, para, cc_list, co_list, path)
                if ok:
                    enviados_fila += 1
                    enviados += 1
                    _log(hoja_mes + "/" + embarque + " | Invoice enviado: " + nombre)
                else:
                    errores += 1
                    _log(hoja_mes + "/" + embarque + " | Falla envio Invoice: " + _err_inv)
            if enviados_fila > 0:
                ws.cell(row=row_idx, column=8, value="ENVIADO")
                from datetime import timezone as _tz, timedelta as _td
                _ts_lima_inv = datetime.now(_tz(_td(hours=-5))).strftime("%d/%m/%Y %H:%M:%S")
                ws.cell(row=row_idx, column=9, value=_ts_lima_inv)

    try:
        wb.save(registro)
        _log("Registro guardado")
    except Exception as e:
        _log("ERROR guardando Registro: " + str(e))
        errores += 1

    # NOTIFICACION DE PENDIENTES - Python SMTP INLINE con HTML
    if pendientes and para_est:
        # Alertas de pendientes van a Estatus
        para = para_est
        cc_list = cc_est
        co_list = co_est
        # Construir tabla HTML (sin columna Carpeta, redundante)
        _MESES_ES = {1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"}
        _rows_html = []
        for (cli_p, emb_p, fec_p, ped_p, dias_p, carpeta_p) in pendientes:
            _cli_h = cli_p.replace("&", "&amp;").replace("<", "&lt;")
            # Carpeta esperada segun fecha: <emb> RIVERSIDE <MES3> <AA>
            _fec_obj_r = None
            for _fmt_r in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                try:
                    from datetime import datetime as _dtr
                    _fec_obj_r = _dtr.strptime(fec_p, _fmt_r).date()
                    break
                except Exception:
                    _fec_obj_r = None
            if _fec_obj_r:
                _car_base = emb_p + " RIVERSIDE " + _MESES_ES[_fec_obj_r.month] + " " + str(_fec_obj_r.year)[-2:]
            else:
                _car_base = emb_p + " RIVERSIDE"
            _car_h = _car_base.replace("&", "&amp;").replace("<", "&lt;")
            # Nombre archivo esperado canonico (PO es opcional, no se muestra)
            _fname = emb_p + " invoice.pdf"
            _fname_h = _fname.replace("&", "&amp;").replace("<", "&lt;")
            _color = "#ffe0e0" if dias_p > 3 else "#fff8e0"
            _rows_html.append(
                "<tr style=\"background:" + _color + ";\">"
                "<td style=\"padding:6px 10px;font-family:monospace;\">" + emb_p + "</td>"
                "<td style=\"padding:6px 10px;\">" + _cli_h + "</td>"
                "<td style=\"padding:6px 10px;\">" + fec_p + "</td>"
                "<td style=\"padding:6px 10px;text-align:center;\"><b>" + str(dias_p) + "</b></td>"
                "<td style=\"padding:6px 10px;font-family:monospace;\">" + ped_p + "</td>"
                "<td style=\"padding:6px 10px;font-family:monospace;font-size:11px;\">" + _car_h + "</td>"
                "<td style=\"padding:6px 10px;font-family:monospace;font-size:11px;\">" + _fname_h + "</td>"
                "</tr>"
            )
        _html = (
            "<html><body style=\"font-family:Calibri,Arial,sans-serif;font-size:13px;color:#222;\">"
            "<p>Estimados,</p>"
            "<p>Se detectaron <b>" + str(len(pendientes)) + " embarques</b> con fecha vencida "
            "cuyo <b>Invoice PDF</b> no se encontro en la carpeta esperada:</p>"
            "<table cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;border:1px solid #999;\">"
            "<thead><tr style=\"background:#333;color:#fff;\">"
            "<th style=\"padding:6px 10px;\">Embarque</th>"
            "<th style=\"padding:6px 10px;\">Cliente</th>"
            "<th style=\"padding:6px 10px;\">Fecha</th>"
            "<th style=\"padding:6px 10px;\">Dias</th>"
            "<th style=\"padding:6px 10px;\">PO</th>"
            "<th style=\"padding:6px 10px;\">Carpeta</th>"
            "<th style=\"padding:6px 10px;\">Nombre archivo esperado</th>"
            "</tr></thead><tbody>"
            + "".join(_rows_html) +
            "</tbody></table>"
            "<p>El bot reintentara el envio en la siguiente corrida.</p>"
            "<p>Saludos,<br><i>Bot COMEX MPF</i></p>"
            "</body></html>"
        )
        _asunto = "COMEX | Invoices pendientes - " + str(hoy)

        # SMTP inline (sin funciones, evita scope issues de Rocketbot)
        import smtplib as _smtp_mod
        from email.mime.text import MIMEText as _MT
        from email.mime.multipart import MIMEMultipart as _MM
        _host_i = GetVar("v_smtp_host") or "smtp.office365.com"
        _port_raw_i = GetVar("v_smtp_port") or "587"
        _port_i = int(_port_raw_i) if (_port_raw_i and _port_raw_i != "ERROR_NOT_VAR") else 587
        _user_i = GetVar("v_smtp_user") if GetVar("v_smtp_user") not in ("", "ERROR_NOT_VAR") else ""
        _pwd_i = GetVar("v_smtp_pass") if GetVar("v_smtp_pass") not in ("", "ERROR_NOT_VAR") else ""
        _ssl_i_raw = GetVar("v_smtp_ssl") or "false"
        _ssl_i = (_ssl_i_raw.lower() == "true") if (_ssl_i_raw and _ssl_i_raw != "ERROR_NOT_VAR") else False

        _log("Enviando correo HTML pendientes via " + _host_i + ":" + str(_port_i) + " user=" + _user_i)
        try:
            _m = _MM("alternative")
            _m["From"] = _user_i
            _m["To"] = ", ".join(para)
            if cc_list:
                _m["Cc"] = ", ".join(cc_list)
            _m["Subject"] = _asunto
            _m.attach(_MT(_html, "html", "utf-8"))
            _destinos_i = para + cc_list + co_list
            if _ssl_i:
                _sm = _smtp_mod.SMTP_SSL(_host_i, _port_i, timeout=30)
            else:
                _sm = _smtp_mod.SMTP(_host_i, _port_i, timeout=30)
                _sm.starttls()
            if _user_i and _pwd_i:
                _sm.login(_user_i, _pwd_i)
            _sm.sendmail(_user_i, _destinos_i, _m.as_string())
            _sm.quit()
            _log("Correo HTML pendientes enviado a " + str(_destinos_i) + " (" + str(len(pendientes)) + " embarques)")
        except Exception as _e_smtp:
            _log("ERROR SMTP pendientes: " + type(_e_smtp).__name__ + ": " + str(_e_smtp))
    elif pendientes and not para_est:
        _log("WARN " + str(len(pendientes)) + " pendientes pero sin destinatarios Estatus/Para.txt")

    detalle = ("total=" + str(total_filas)
              + " ya_enviado=" + str(ya_enviado)
              + " futuro=" + str(futuro)
              + " mismo_dia=" + str(mismo_dia)
              + " procesados=" + str(procesados)
              + " enviados=" + str(enviados)
              + " sin_archivo=" + str(sin_archivo)
              + " atrasados=" + str(atrasados)
              + " errores=" + str(errores))
    SetVar("v_h4_estado", "OK" if errores == 0 else "OK_CON_ERRORES")
    SetVar("v_h4_detalle", detalle)
    _log("Resumen | " + detalle)

_log("Fin bot 4_Revision_Invoice")
