from datetime import datetime, date
import os

hoy = datetime.now()
iso_anio = hoy.isocalendar()[0]

root_embarques = GetVar("v_ruta_base")
ruta_anio = os.path.join(root_embarques, str(iso_anio))

SetVar("v_anio_semana", str(iso_anio))
SetVar("v_ruta_anio", ruta_anio)
SetVar("v_h1_estado", "INICIADO")
SetVar("v_h1_detalle", "Preparando estructura Año > Semana del año")
