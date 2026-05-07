from datetime import datetime
import os as _os
rp = GetVar("v_orq_log_path")
if rp and rp != "ERROR_NOT_VAR" and _os.path.isabs(rp):
    try:
        detalle = "archivo=" + str(GetVar("v_ruta_excel")) + " | filas=" + str(GetVar("v_total_filas"))
        with open(rp, "a", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|2_Embarques " + str(GetVar("v_orq_estado_2")) + " | " + detalle + chr(10))
    except Exception:
        pass
