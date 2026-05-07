from datetime import datetime
inicio_txt = GetVar("v_orq_inicio")
fmt = "%Y-%m-%d %H:%M:%S"
fin = datetime.now()
SetVar("v_orq_fin", fin.strftime(fmt))
dur = ""
try:
    if inicio_txt:
        dur = str(fin - datetime.strptime(inicio_txt, fmt))
except Exception:
    dur = ""
SetVar("v_orq_duracion", dur)
ruta_log = GetVar("v_orq_log_path")
if ruta_log:
    with open(ruta_log, "a", encoding="utf-8") as f:
        f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|Fin del flujo | estado=" + str(GetVar("v_orq_estado_general")) + " | duracion=" + dur + chr(10))
