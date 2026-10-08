@echo off
rem Abre el programa sin ventana de consola
cd /d "%~dp0"
where pythonw >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python. Instalelo desde https://www.python.org/downloads/
  echo y marque la opcion "Add python.exe to PATH" durante la instalacion.
  pause
  exit /b 1
)
start "" pythonw "%~dp0contabilidad.pyw"
