@echo off
chcp 65001 >nul
title Subir catalogo a GitHub
cd /d "%~dp0"

echo ========================================================
echo   SUBIENDO PROYECTO A GITHUB
echo ========================================================
echo.
echo 1. Verificando cambios locales...
git add .
git commit -m "Actualizacion de catalogo" >nul 2>&1

echo 2. Enviando a GitHub (origin/main)...
echo    (Si se abre una ventana del navegador, autoriza tu cuenta)
echo.
git push -u origin main

if %ERRORLEVEL% equ 0 (
    echo.
    echo ========================================================
    echo   ¡EXITO! El proyecto se subio correctamente a GitHub.
    echo ========================================================
) else (
    echo.
    echo [!] Hubo un detalle al subir. Revisa el mensaje arriba o autoriza tu cuenta.
)
echo.
pause
