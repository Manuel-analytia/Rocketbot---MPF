from datetime import datetime
import os as _os
rp = GetVar("v_orq_log_path")
if rp and rp != "ERROR_NOT_VAR" and _os.path.isabs(rp):
    try:
        detalle = "No se encontro archivo Excel o no habia filas para continuar con los hijos 3 y 4"
        with open(rp, "a", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|2_Embarques " + str(GetVar("v_orq_estado_2")) + " | " + detalle + chr(10))
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|3_Revision_COAS_Informe OMITIDO" + chr(10))
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|4_Revision_Invoice OMITIDO" + chr(10))
    except Exception:
        pass
