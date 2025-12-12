@echo off
REM Script para iniciar el servidor de desarrollo de Nexus Orchestrator
REM sin necesidad de tener Node.js instalado globalmente

echo ========================================
echo   Nexus Orchestrator - Dev Server
echo ========================================
echo.

REM Configurar PATH temporal para esta sesión
set "NODE_PATH=%~dp0..\node-v24.11.1-win-x64"
set "PATH=%NODE_PATH%;%PATH%"

REM Verificar que Node.js está disponible
echo [1/3] Verificando Node.js...
node --version
if %errorlevel% neq 0 (
    echo ERROR: No se pudo encontrar Node.js
    echo Asegurate de que la carpeta node-v24.11.1-win-x64 existe
    pause
    exit /b 1
)

echo [2/3] Verificando npm...
npm.cmd --version
if %errorlevel% neq 0 (
    echo ERROR: No se pudo encontrar npm
    pause
    exit /b 1
)

echo.
echo [3/3] Iniciando servidor de desarrollo...
echo.
echo La aplicacion estara disponible en: http://localhost:3000
echo Presiona Ctrl+C para detener el servidor
echo.

REM Iniciar el servidor de desarrollo
npm.cmd run dev

pause
