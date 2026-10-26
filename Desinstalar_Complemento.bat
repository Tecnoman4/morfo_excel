@echo off
title Desinstalador - Morfo para Excel
cls
echo ==============================================================
echo             DESINSTALADOR DE MORFO PARA EXCEL
echo ==============================================================
echo.
echo Desinstalando complemento Morfo de Microsoft Excel...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Uninstall-Addin.ps1"

echo.
pause
