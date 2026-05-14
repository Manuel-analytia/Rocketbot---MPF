@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM  bot_comex.bat - Lanzador del Bot RPA - Envio de correos COMEX
REM  Diseñado para correr desatendido desde el Programador de Tareas de Windows.
REM  En servidor se ubica en C:\Scripts\bot_comex.bat (ajustable).
REM ============================================================================
REM
REM  Para PRUEBAS LOCALES (OneDrive ya corriendo, sync ya hecho):
REM      bajar SYNC_WAIT_SEC a 5 o 10. No es necesario tocar nada mas.
REM
REM  Para PRODUCCION en SERVIDOR (sesion bloqueada, sync de varias horas):
REM      mantener SYNC_WAIT_SEC en 120-240 segundos.
REM ============================================================================

REM ===== CONFIG (unico bloque editable) =======================================
set ROCKETBOT_DIR=C:\Program Files (x86)\Rocketbot
set ROCKETBOT_EXE=rocketbot.exe
set ROBOT_NAME=0_Flujo_Regular
set CLI_FLAG=-start=
set SYNC_WAIT_SEC=120
set BAT_LOG_DIR=C:\Scripts\logs
REM ============================================================================

REM ===== Crear directorio de log si no existe =================================
if not exist "%BAT_LOG_DIR%" mkdir "%BAT_LOG_DIR%" 2>nul
set BAT_LOG=%BAT_LOG_DIR%\bot_comex_bat.log

echo. >> "%BAT_LOG%"
echo ===== %DATE% %TIME% ===== >> "%BAT_LOG%"
echo [INICIO] Usuario=%USERNAME%  Maquina=%COMPUTERNAME% >> "%BAT_LOG%"

REM ===== 1. Lanzar OneDrive en background si no esta corriendo ================
tasklist /FI "IMAGENAME eq OneDrive.exe" 2>nul | find /I "OneDrive.exe" >nul
if errorlevel 1 (
    echo [%TIME%] OneDrive no esta corriendo, lanzandolo... >> "%BAT_LOG%"
    start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /background
) else (
    echo [%TIME%] OneDrive ya estaba corriendo >> "%BAT_LOG%"
)

REM ===== 2. Pausa para que sincronice =========================================
echo [%TIME%] Esperando %SYNC_WAIT_SEC%s para sync de OneDrive... >> "%BAT_LOG%"
timeout /t %SYNC_WAIT_SEC% /nobreak >nul

REM ===== 3. Validar que Rocketbot existe ======================================
if not exist "%ROCKETBOT_DIR%\%ROCKETBOT_EXE%" (
    echo [ERROR] No se encontro %ROCKETBOT_EXE% en %ROCKETBOT_DIR% >> "%BAT_LOG%"
    echo [ERROR] Editar la variable ROCKETBOT_DIR en este .bat con la ruta real. >> "%BAT_LOG%"
    exit /b 1
)

REM ===== 4. Ejecutar Rocketbot ================================================
echo [%TIME%] Ejecutando: %ROCKETBOT_EXE% %CLI_FLAG% "%ROBOT_NAME%" >> "%BAT_LOG%"
cd /d "%ROCKETBOT_DIR%"
start "" /WAIT "%ROCKETBOT_EXE%" %CLI_FLAG%%ROBOT_NAME%
set RB_EXIT=%ERRORLEVEL%

echo [%TIME%] Rocketbot termino con exit code %RB_EXIT% >> "%BAT_LOG%"
echo ===== FIN ===== >> "%BAT_LOG%"

exit /b %RB_EXIT%
