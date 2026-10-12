@echo off
title Actualizar Datos SAP - Portal de Traslados
echo ==========================================================
echo   Actualizando Datos desde ME2W.xlsx y ME80FN.xlsx
echo ==========================================================
powershell -ExecutionPolicy Bypass -File "%~dp0actualizar_datos.ps1"
echo.
echo Presiona cualquier tecla para salir...
pause >nul
