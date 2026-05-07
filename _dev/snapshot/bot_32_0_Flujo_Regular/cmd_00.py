import os
from datetime import datetime

# Escribir trace inline (sin helper _tr porque Rocketbot no ve variables
# modulo desde funciones).
_HOME = os.path.expanduser("~")
_OD = None
_marker_sub = "SAN ISIDRO - PROGRAMACION SEMANAL DE EMBARQUES"
try:
    for _nn in os.listdir(_HOME):
        if _nn.lower().startswith("onedrive"):
            _cand = os.path.join(_HOME, _nn)
            if os.path.isdir(os.path.join(_cand, _marker_sub)):
                _OD = _cand
                break
except Exception:
    pass
if not _OD:
    _OD = os.path.join(_HOME, "OneDrive - mpf.com.pe (1)")
    if not os.path.isdir(_OD):
        _OD = os.path.join(_HOME, "OneDrive - mpf.com.pe")
_T = os.path.join(_OD, "rocketbot_trace.txt")

usuario = os.environ.get("USERNAME", "")
_isdir = os.path.isdir(_OD)

try:
    with open(_T, "a", encoding="utf-8") as f:
        f.write(str(datetime.now()) + " | CMD00 INICIO | USERNAME=" + repr(usuario)
                + " home=" + repr(_HOME) + " ruta=" + repr(_OD)
                + " isdir=" + str(_isdir) + chr(10))
except Exception:
    pass

SetVar("v_root_onedrive", _OD)

try:
    with open(_T, "a", encoding="utf-8") as f:
        f.write(str(datetime.now()) + " | CMD00 FIN SetVar v_root_onedrive=" + _OD + chr(10))
except Exception:
    pass
