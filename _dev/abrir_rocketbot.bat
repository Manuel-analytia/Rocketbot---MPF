@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM  abrir_rocketbot.bat - Lanzador de desarrollo
REM ============================================================================
REM  Abre Rocketbot Studio leyendo el robot.db del repo en lugar del default
REM  en C:\Program Files (x86)\Rocketbot\robot.db. Lo logra creando un HARD
REM  LINK NTFS, asi cualquier cambio guardado desde Rocketbot Studio queda
REM  reflejado en el repo automaticamente y se puede commitear.
REM
REM  - Cierra Rocketbot Studio antes de correr este .bat.
REM  - Auto-eleva con UAC porque escribir en Program Files requiere admin.
REM  - Deja la ventana abierta al terminar para revisar el log.
REM
REM  IMPORTANTE: dentro de bloques if (...) se usa !VAR! en lugar de %VAR%
REM  porque las rutas tienen "(x86)" y los parentesis rompen el parseo de cmd
REM  con expansion temprana.
REM ============================================================================

REM ===== Auto-elevacion a administrador =======================================
fltmc >nul 2>&1
if %errorlevel% NEQ 0 (
    echo Solicitando elevacion a administrador...
    powershell -Command "Start-Process cmd.exe -ArgumentList '/K \"%~f0\"' -Verb RunAs"
    exit /b
)

REM ===== CONFIG ===============================================================
set REPO_DB=C:\Repositorios\MPF - Rocketbot\Rocketbot---MPF\robot.db
set ROCKETBOT_DIR=C:\Program Files (x86)\Rocketbot
set RB_DB=%ROCKETBOT_DIR%\robot.db
set ROCKETBOT_EXE=rocketbot.exe
REM ============================================================================

echo.
echo ===== abrir_rocketbot.bat (admin OK) =====
echo Repo db   : %REPO_DB%
echo Rocketbot : %RB_DB%
echo.

REM ===== 1. Validar repo db existe ============================================
echo [1/5] Validando repo db...
if not exist "!REPO_DB!" (
    echo   [ERROR] No existe el robot.db del repo: "!REPO_DB!"
    pause
    exit /b 1
)
echo   OK

REM ===== 2. Validar Rocketbot instalado =======================================
echo [2/5] Validando Rocketbot...
if not exist "!ROCKETBOT_DIR!\!ROCKETBOT_EXE!" (
    echo   [ERROR] No se encontro "!ROCKETBOT_EXE!" en "!ROCKETBOT_DIR!"
    pause
    exit /b 1
)
echo   OK

REM ===== 3. Detectar si ya hay hard link al repo ==============================
echo [3/5] Verificando hard link...
set ALREADY_LINKED=0
if exist "!RB_DB!" (
    fsutil hardlink list "!RB_DB!" 2>nul | findstr /I /C:"Rocketbot---MPF" >nul
    if not errorlevel 1 set ALREADY_LINKED=1
)

if "!ALREADY_LINKED!"=="1" (
    echo   El hard link ya existe.
    goto :LAUNCH
)
echo   No hay link aun, lo voy a crear.

REM ===== 4. Backup del robot.db actual de Rocketbot (una sola vez) ===========
echo [4/5] Backup y reemplazo del robot.db default...
if exist "!RB_DB!" (
    if not exist "!RB_DB!.backup-original" (
        echo   Backup: "!RB_DB!" -^> "!RB_DB!.backup-original"
        copy /Y "!RB_DB!" "!RB_DB!.backup-original" >nul
        if errorlevel 1 (
            echo   [ERROR] No se pudo hacer backup. Permisos?
            pause
            exit /b 1
        )
    ) else (
        echo   Backup ya existe en "!RB_DB!.backup-original"
    )

    echo   Eliminando "!RB_DB!" actual...
    del /F /Q "!RB_DB!" 2>nul
    if exist "!RB_DB!" (
        echo   [ERROR] No se pudo borrar "!RB_DB!"
        echo           ^>^> Probablemente Rocketbot Studio esta abierto. Cierralo y reintenta.
        pause
        exit /b 1
    )
    echo   Borrado OK.
)

REM ===== 5. Crear hard link ===================================================
echo [5/5] Creando hard link...
mklink /H "!RB_DB!" "!REPO_DB!"
if errorlevel 1 (
    echo   [WARN] mklink fallo. Como fallback voy a COPIAR el del repo.
    echo          OJO: cambios hechos en Rocketbot NO se reflejaran en el repo.
    copy /Y "!REPO_DB!" "!RB_DB!"
    if errorlevel 1 (
        echo   [ERROR] Tampoco se pudo copiar.
        pause
        exit /b 1
    )
    echo   Copia OK ^(sin link, hay que copiar manual al cerrar Rocketbot^).
) else (
    echo   Hard link OK.
)

:LAUNCH
echo.
echo ===== Lanzando Rocketbot Studio... =====
cd /d "!ROCKETBOT_DIR!"
start "" "!ROCKETBOT_EXE!"

echo.
echo Listo. Si todo salio bien, deberia aparecer Rocketbot Studio con los 6 bots:
echo   - 0_Flujo_Regular (orquestador)
echo   - 0_1_Verificacion_Condiciones
echo   - 1_Crear_Carpetas
echo   - 2_Embarques
echo   - 3_Revision_COAS_Informe
echo   - 4_Revision_Invoice
echo.
echo Esta ventana se queda abierta para revisar el log. Cierrala cuando termines.
echo.

endlocal
