@echo off
chcp 65001 >nul
title 🚀 Nexus Jarvis Automation Suite - CIAL

echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║     🤖 SAP AUTOMATION SUITE - CIAL ALIMENTOS 🤖           ║
echo ╚════════════════════════════════════════════════════════════╝
echo.

:: Obtener la ruta del directorio donde está el .bat
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo 📂 Directorio de trabajo: %SCRIPT_DIR%
echo.

:: --- CONFIGURACIÓN NODE LOCAL ---
set "NODE_HOME=%SCRIPT_DIR%node-v24.11.1-win-x64"
set "PATH=%NODE_HOME%;%PATH%"

echo [1/4] 🔍 Verificando Entorno...
echo    -> Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python NO está instalado.
    echo 💡 Ejecuta "installer.bat" primero.
    pause
    exit /b 1
)
echo    -> Node.js (Local)...
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Error configurando Node.js local.
    pause
    exit /b 1
)
echo ✅ Entorno OK
echo.

:: Iniciar procesos (Worker y Web Server)
echo [2/4] 🚀 Iniciando Servicios...

:: Siempre usar ventanas separadas (más confiable)
echo    -> Iniciando Jarvis Worker SAP...
start "Jarvis Worker SAP" cmd /k "python worker_sap.py"

echo    -> Iniciando Consolidación Zonales...
start "Consolidación Zonales" cmd /k "python Bot_Consolidacion_Zonales.py"

echo    -> Iniciando Servidor Web...
cd Interfaz
start "Nexus Web Server" cmd /k "set PATH=%NODE_HOME%;%PATH% && npm run dev"
cd ..

:: Esperar a que Vite inicie (aprox 5 seg)
echo ⏳ Cargando interfaz...
timeout /t 10 /nobreak >nul

echo [4/4] 🖥️ Verificando Navegador...
powershell -Command "$t='Nexus Orchestrator'; $w=Get-Process | Where-Object {$_.MainWindowTitle -match $t}; if ($w) { Write-Host '   -> Ya está abierto. Saltando apertura.' } else { Start-Process 'http://localhost:3000' }"

echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║                  ✅ SISTEMA INICIADO ✅                    ║
echo ╠════════════════════════════════════════════════════════════╣
echo ║                                                            ║
echo ║  1. Jarvis Worker SAP: Ejecutándose (Ventana Negra)       ║
echo ║  2. Consolidación Zonales: Ejecutándose (Ventana Negra)   ║
echo ║  3. Servidor Web: Ejecutándose (Minimizado)               ║
echo ║  4. Interfaz: http://localhost:3000                       ║
echo ║                                                            ║
echo ║  ⚠️  NO CIERRES las ventanas negras                       ║
echo ║                                                            ║
echo ╚════════════════════════════════════════════════════════════╝
echo.
echo 💡 Para detener, cierra todas las ventanas.
echo.

pause >nul
