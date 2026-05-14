@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM  ejecutar_bot.bat - Ejecuta el orquestador del bot via CLI de Rocketbot.
REM
REM  DEPLOY: copiar este archivo dentro de la carpeta donde esta rocketbot.exe
REM  (en este caso C:\Program Files (x86)\Rocketbot\). El .bat resuelve su
REM  propia ruta con %~dp0, asi no hay paths hardcodeados.
REM
REM  Requiere Production License activada en Rocketbot.
REM  Auto-elevacion COMENTADA por ahora (descomentar el bloque de abajo si
REM  hace falta admin para escribir en Program Files o para drivers).
REM ============================================================================

REM ===== Auto-elevacion a administrador (DESACTIVADO) =========================
REM fltmc >nul 2>&1
REM if %errorlevel% NEQ 0 (
REM     echo Solicitando elevacion a administrador...
REM     powershell -Command "Start-Process cmd.exe -ArgumentList '/K \"%~f0\"' -Verb RunAs"
REM     exit /b
REM )

REM ===== Resolver paths desde la ubicacion del .bat ===========================
REM %~dp0 termina con barra; lo guardamos sin barra final tambien por si acaso
set HERE=%~dp0
set HERE_NB=%HERE:~0,-1%
set ROBOT_NAME=0_Flujo_Regular

echo.
echo ===== ejecutar_bot.bat (admin OK) =====
echo Carpeta   : !HERE_NB!
echo Robot     : !ROBOT_NAME!
echo.

REM ===== Validar rocketbot.exe esta en la misma carpeta =======================
if not exist "!HERE!rocketbot.exe" (
    echo [ERROR] No se encontro rocketbot.exe en "!HERE_NB!"
    echo Copiar este .bat dentro de la carpeta de instalacion de Rocketbot.
    pause
    exit /b 1
)

REM ===== Ejecutar Rocketbot ===================================================
cd /d "!HERE_NB!"
rocketbot.exe -start=!ROBOT_NAME!
set RB_EXIT=!ERRORLEVEL!

echo.
echo Rocketbot termino con exit code !RB_EXIT!
pause
endlocal
exit /b %RB_EXIT%
