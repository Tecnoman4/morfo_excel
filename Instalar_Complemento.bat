@echo off
title Instalador - Morfo para Excel
cls
echo ==============================================================
echo              INSTALADOR DE MORFO PARA EXCEL
echo ==============================================================
echo.
echo Instalando complemento Morfo en Microsoft Excel...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Install-Addin.ps1"

echo.
pause
