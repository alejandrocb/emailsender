@echo off
chcp 65001 >nul
title SCS - Instalar Pedidos Liberados
echo ============================================================
echo   SCS - PEDIDOS LIBERADOS - INSTALACION SIN .NET SDK
echo ============================================================
echo.
net session >nul 2>&1
if not "%errorlevel%"=="0" (
  echo [ERROR] Debe ejecutar este archivo con boton derecho:
  echo         "Ejecutar como administrador".
  echo.
  pause
  exit /b 1
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Instalar-Tarea.ps1"
echo.
pause
