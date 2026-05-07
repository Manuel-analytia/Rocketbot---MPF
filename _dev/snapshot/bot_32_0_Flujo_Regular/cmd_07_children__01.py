from datetime import datetime
import os as _os
rp = GetVar("v_orq_log_path")
if rp and rp != "ERROR_NOT_VAR" and _os.path.isabs(rp):
    try:
        with open(rp, "a", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|Iniciando 1_Crear_Carpetas" + chr(10))
    except Exception:
        pass
