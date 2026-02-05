@echo off
:: Navegar a la carpeta del script usando comillas para evitar errores por espacios
cd /d "%~dp0"

:: Definir la ruta al ejecutable de Python embebido (con comillas)
set "PYTHON_EXE=%~dp0python-3.13-embed\python.exe"

echo Iniciando el bot de SAP...

:: Ejecutar el bot. Las comillas en "%PYTHON_EXE%" son la clave.
"%PYTHON_EXE%" main.py

:: Verificamos si hubo error
if %ERRORLEVEL% neq 0 (
    echo.
    echo Ocurrio un error al ejecutar el bot.
    pause
)