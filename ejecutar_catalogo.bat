@echo off
title Generador de Catalogo - Palacio de la Limpieza

cd /d "%~dp0"

echo ============================================
echo   Generador de Catalogo - Palacio
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta agregado al PATH.
    echo Instala Python 3 y marca la opcion "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

echo Generando catalogo...
python generar_catalogo.py

echo.
if errorlevel 1 (
    echo.
    echo Ocurrio un error al generar el catalogo.
    echo.
    pause
    exit /b 1
) else (
    start "" "salida\catalogo.pdf"
    exit /b 0
)
