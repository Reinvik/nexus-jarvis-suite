@echo off
REM Script para instalar dependencias de Nexus Orchestrator
REM sin necesidad de tener Node.js instalado globalmente

echo ========================================
echo   Nexus Orchestrator - Instalador
echo ========================================
echo.

REM Configurar PATH temporal para esta sesión
set "NODE_PATH=%~dp0..\node-v24.11.1-win-x64"
set "PATH=%NODE_PATH%;%PATH%"

REM Verificar que Node.js está disponible
echo [1/2] Verificando Node.js...
node --version
if %errorlevel% neq 0 (
    echo ERROR: No se pudo encontrar Node.js
    echo Asegurate de que la carpeta node-v24.11.1-win-x64 existe
    pause
    exit /b 1
)

echo [2/2] Instalando dependencias...
echo Esto puede tomar varios minutos...
echo.

npm.cmd install

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo   Instalacion completada exitosamente!
    echo ========================================
    echo.
    echo Ahora puedes ejecutar: iniciar-dev.bat
    echo.
) else (
    echo.
    echo ERROR: La instalacion fallo
    echo.
)

pause
