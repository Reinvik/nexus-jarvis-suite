@echo off
echo ========================================
echo   REINICIANDO NEXUS ORCHESTRATOR
echo ========================================
echo.

REM Detener procesos existentes
echo [1/4] Deteniendo procesos anteriores...
taskkill /F /IM python.exe /FI "WINDOWTITLE eq Worker SAP*" 2>nul
taskkill /F /IM node.exe /FI "WINDOWTITLE eq Interfaz Web*" 2>nul
timeout /t 2 /nobreak >nul

REM Reconstruir interfaz
echo.
echo [2/4] Reconstruyendo interfaz web...
cd /d "%~dp0Interfaz"
call npm run build
if errorlevel 1 (
    echo ERROR: No se pudo construir la interfaz
    pause
    exit /b 1
)

REM Volver al directorio principal
cd /d "%~dp0"

REM Iniciar worker
echo.
echo [3/4] Iniciando Worker SAP...
start "Worker SAP - SANJORGE1" cmd /k "python worker_sap.py"
timeout /t 3 /nobreak >nul

REM Iniciar interfaz
echo.
echo [4/4] Iniciando Interfaz Web...
cd /d "%~dp0Interfaz"
start "Interfaz Web - Nexus" cmd /k "npm run dev"

REM Esperar y abrir navegador
timeout /t 5 /nobreak >nul
start https://nexus-orchestrator.vercel.app/

echo.
echo ========================================
echo   SISTEMA REINICIADO CORRECTAMENTE
echo ========================================
echo.
echo Worker SAP: Activo
echo Interfaz Web: Activo
echo.
pause
