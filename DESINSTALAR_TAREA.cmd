@echo off
chcp 65001 >nul
net session >nul 2>&1
if not "%errorlevel%"=="0" (
  echo Ejecute como administrador.
  pause
  exit /b 1
)
schtasks /Delete /TN "SCS - Pedidos Liberados" /F
pause
