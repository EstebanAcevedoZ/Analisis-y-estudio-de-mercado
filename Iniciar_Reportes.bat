@echo off
title Servidor RenoBo
echo ====================================================
echo    Iniciando Generador de Reportes de Mercado...
echo    Por favor NO cierres esta ventana negra.
echo ====================================================
echo.

:: Esto fuerza a Windows a ubicarse en la carpeta donde está este archivo
cd /d "%~dp0"

:: Ejecuta la aplicación
python app.py

pause