from datetime import datetime
ruta_log = GetVar("v_orq_log_path")
if ruta_log and ruta_log != "ERROR_NOT_VAR":
    try:
        with open(ruta_log, "a", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|Inicio del flujo orquestador @ " + str(GetVar("v_orq_inicio")) + chr(10))
            f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|Precheck " + str(GetVar("v_precheck_estado")) + " | " + str(GetVar("v_precheck_detalle")) + chr(10))
    except Exception:
        pass
