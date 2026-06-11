@echo off
REM ===============================
REM Comprobar si las librerías necesarias están instaladas
REM ===============================

REM Lista de librerías necesarias
set LIBRERIAS=selenium requests pandas openpyxl

REM Bandera para verificar si hay que instalar
set FALTANTES=0

REM Activar expansión retardada de variables
setlocal enabledelayedexpansion

REM Verificar si las librerías están instaladas
for %%i in (%LIBRERIAS%) do (
    pip show %%i >nul 2>&1
    if errorlevel 1 (
        echo La libreria %%i no esta instalada.
        set /a FALTANTES+=1
    ) else (
        echo La libreria %%i ya esta instalada.
    )
)

REM Si hay alguna librería faltante, instalar todas desde el archivo
if !FALTANTES! GTR 0 (
    echo.
    echo Instalando librerias desde carpeta local...
    pip install --no-index --find-links=LIbs -r requirements.txt
)

REM Limpiar pantalla
cls
echo ===============================
echo Se inicia la ejecucion de la RPA
echo ===============================

REM Ejecutar el script de Python con la interfaz Tkinter
python App/gui.py


echo.
echo ===== FIN DEL PROCESO =====
echo.
pause


