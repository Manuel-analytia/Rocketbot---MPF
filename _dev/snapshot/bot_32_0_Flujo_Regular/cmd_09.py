from datetime import datetime
import os
rp = GetVar("v_orq_log_path")
if rp and rp != "ERROR_NOT_VAR" and os.path.isabs(rp):
    try:
        msg = ("Resumen orquestador: "
               + str(GetVar("v_orq_estado_general"))
               + " | PRECHECK=" + str(GetVar("v_orq_estado_0_1"))
               + " | H1=" + str(GetVar("v_orq_estado_1"))
               + " | H2=" + str(GetVar("v_orq_estado_2"))
               + " | H3=" + str(GetVar("v_orq_estado_3"))
               + " | H4=" + str(GetVar("v_orq_estado_4")))
        with open(rp, "a", encoding="utf-8") as _f:
            _f.write(datetime.now().strftime("%H:%M:%S") + "|0_Flujo_Regular|" + msg + chr(10))
    except Exception:
        pass
