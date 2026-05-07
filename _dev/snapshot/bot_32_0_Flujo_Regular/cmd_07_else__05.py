from datetime import datetime
import os as _os
rp = GetVar("v_orq_log_path")
if rp and rp != "ERROR_NOT_VAR" and _os.path.isabs(rp):
    try:
        with open(rp, "a", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|Flujo detenido por precheck | " + str(GetVar("v_precheck_detalle")) + chr(10))
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|1_Crear_Carpetas DETENIDO_PRECHECK OMITIDO" + chr(10))
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|2_Embarques DETENIDO_PRECHECK OMITIDO" + chr(10))
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|3_Revision_COAS_Informe DETENIDO_PRECHECK OMITIDO" + chr(10))
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|4_Revision_Invoice DETENIDO_PRECHECK OMITIDO" + chr(10))
    except Exception:
        pass
