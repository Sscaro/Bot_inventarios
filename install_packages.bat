@echo off
:: Navegar a la carpeta donde reside este script
cd /d "%~dp0"

echo --- Iniciando instalacion de dependencias ---

:: Definir la ruta al Python embebido
set PYTHON_EXE="%~dp0python-3.13-embed\python.exe"

:: Ejecutar la instalacion
%PYTHON_EXE% -m pip install -r requirements.txt

echo.
echo --- Proceso finalizado ---
pause