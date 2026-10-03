@echo off
title Sincronizar Fotos con Excel - Palacio de la Limpieza
cd /d "%~dp0"

echo ===================================================
echo   Sincronizar Fotos Nuevas con catalogo_productos.xlsx
echo ===================================================
echo.

python sincronizar_fotos.py

echo.
pause
