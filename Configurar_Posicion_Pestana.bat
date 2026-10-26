@echo off
title Posicion de la Pestana - Morfo para Excel
cls
echo ==============================================================
echo        CONFIGURAR POSICION DE LA PESTANA MORFO EN EXCEL
echo ==============================================================
echo.
echo Selecciona donde deseas ubicar la pestana 'Morfo' en la cinta:
echo.
echo   [1] Al final de la cinta de opciones (Por defecto / Recomendado)
echo   [2] Despues de Inicio (Home)
echo   [3] Despues de Datos (Data)
echo   [4] Despues de Revisar (Review)
echo   [5] Despues de Vista (View)
echo   [6] Despues de Programador / Desarrollador (Developer)
echo.
set /p opt="Elige una opcion (1-6) [Enter para 1]: "

set pos=Final
if "%opt%"=="1" set pos=Final
if "%opt%"=="2" set pos=Inicio
if "%opt%"=="3" set pos=Datos
if "%opt%"=="4" set pos=Revisar
if "%opt%"=="5" set pos=Vista
if "%opt%"=="6" set pos=Programador

echo.
echo Aplicando posicion: %pos%...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Install-Addin.ps1" -TabPosition %pos%
echo.
pause
