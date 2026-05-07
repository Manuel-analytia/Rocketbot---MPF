from datetime import datetime
import os as _os
rp = GetVar("v_orq_log_path")
if rp and rp != "ERROR_NOT_VAR" and _os.path.isabs(rp):
    try:
        estado = str(GetVar("v_orq_estado_3"))
        detalle = str(GetVar("v_h3_detalle"))
        linea = datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|3_Revision_COAS_Informe " + estado
        if detalle:
            linea = linea + " | " + detalle
        with open(rp, "a", encoding="utf-8") as f:
            f.write(linea + chr(10))
    except Exception:
        pass
