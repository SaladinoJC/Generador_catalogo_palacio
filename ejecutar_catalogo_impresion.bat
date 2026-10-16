@echo off
title Catalogo para imprimir - Palacio de la Limpieza

cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta agregado al PATH.
    pause
    exit /b 1
)

echo Generando catalogo para imprimir, sin precios ni botones...
python generar_catalogo.py --imprimir
if errorlevel 1 (
    echo Ocurrio un error al generar el catalogo para imprimir.
    pause
    exit /b 1
)

start "" "salida\catalogo_impresion.pdf"
exit /b 0
