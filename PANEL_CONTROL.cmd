@echo off
chcp 65001 >nul
set "PANEL=\\gerencialz.canariasalud\archivos\Logistica\COMPRAS\Pedidos Liberados\reportes\panel_control.html"
if not exist "%PANEL%" (
  echo El panel aun no existe. Ejecute primero EJECUTAR_AHORA.cmd.
  pause
  exit /b 1
)
start "" "%PANEL%"
