@echo off
setlocal
chcp 65001 >nul
title Cortes de obra - Darca II
echo.
echo Iniciando cortes de obra Darca II...
echo Carpeta: %~dp0
echo Espere mientras se comprueba Python.
echo.
set "DARCA_SIN_PAUSA="
for %%A in (%*) do if /I "%%~A"=="-SinPausa" set "DARCA_SIN_PAUSA=1"
if not exist "%~dp0iniciar_corte.ps1" goto archivos_faltantes
if not exist "%~dp0generar_corte.py" goto archivos_faltantes
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0iniciar_corte.ps1" %*
set "DARCA_RESULTADO=%errorlevel%"
if "%DARCA_RESULTADO%"=="0" exit /b 0
echo.
echo El programa termino con codigo %DARCA_RESULTADO%. Copie el mensaje anterior.
if not defined DARCA_SIN_PAUSA pause
exit /b %DARCA_RESULTADO%

:archivos_faltantes
echo ERROR: faltan archivos del programa junto a INICIAR_CORTE.bat.
echo Extraiga TODO el ZIP y abra el iniciador desde la carpeta extraida.
echo No ejecute este archivo desde el ZIP ni lo copie solo a otra carpeta.
if not defined DARCA_SIN_PAUSA pause
exit /b 2
