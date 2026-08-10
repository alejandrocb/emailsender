@echo off
chcp 65001 >nul
echo ============================================================
echo   SCS - PEDIDOS LIBERADOS - EJECUCION MANUAL
echo ============================================================
echo.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0PedidosLiberados.ps1" -Manual
set RC=%ERRORLEVEL%
echo.
if "%RC%"=="0" (
  echo [OK] Revision terminada correctamente.
  echo Consulte la carpeta reportes.
) else (
  echo [ERROR] La ejecucion ha fallado. Codigo de salida: %RC%
  echo Revise tambien: %%ProgramData%%\SCS\PedidosLiberados
)
echo.
pause
exit /b %RC%
