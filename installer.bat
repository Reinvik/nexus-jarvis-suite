@echo off
chcp 65001 >nul
title 🔧 Instalador - SAP Automation Suite

echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║     🔧 INSTALADOR - SAP AUTOMATION SUITE 🔧               ║
echo ╚════════════════════════════════════════════════════════════╝
echo.
echo Este script instalará todas las dependencias necesarias.
echo Solo necesitas ejecutarlo UNA VEZ.
echo.
pause

:: Obtener la ruta del directorio donde está el .bat
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo.
echo [1/3] 🔍 Verificando Python...
echo.

:: Verificar si Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python NO está instalado en este equipo.
    echo.
    echo ┌─────────────────────────────────────────────────────────┐
    echo │  📥 DESCARGA PYTHON MANUALMENTE:                        │
    echo │                                                         │
    echo │  1. Ve a: https://www.python.org/downloads/             │
    echo │  2. Descarga Python 3.11 o superior                     │
    echo │  3. Durante la instalación:                             │
    echo │     ✅ Marca "Add Python to PATH"                       │
    echo │     ✅ Instala para todos los usuarios                  │
    echo │  4. Reinicia esta ventana después de instalar           │
    echo └─────────────────────────────────────────────────────────┘
    echo.
    pause
    exit /b 1
)

python --version
echo ✅ Python está instalado correctamente
echo.

:: Actualizar pip
echo [2/3] 📦 Actualizando pip...
python -m pip install --upgrade pip
echo.

:: Instalar dependencias
echo [3/3] 📦 Instalando dependencias...
echo.

if exist "requirements.txt" (
    echo 📄 Instalando desde requirements.txt...
    python -m pip install -r requirements.txt
) else (
    echo ⚠️  No se encontró requirements.txt
    echo 📦 Instalando dependencias básicas...
    
    python -m pip install firebase-admin
    python -m pip install customtkinter
    python -m pip install openpyxl
    python -m pip install pandas
    python -m pip install pywin32
    python -m pip install requests
    
    echo.
    echo ℹ️  Si necesitas más dependencias, créalas en requirements.txt
)

echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║              ✅ INSTALACIÓN COMPLETADA ✅                  ║
echo ╠════════════════════════════════════════════════════════════╣
echo ║                                                            ║
echo ║  🎉 Todo está listo para usar el sistema                  ║
echo ║                                                            ║
echo ║  📌 Próximo paso:                                          ║
echo ║     Ejecuta "launcher.bat" para iniciar el sistema        ║
echo ║                                                            ║
echo ╚════════════════════════════════════════════════════════════╝
echo.
pause
