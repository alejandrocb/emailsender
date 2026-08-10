@echo off
chcp 65001 >nul
echo ============================================================
echo DIAGNOSTICO - PEDIDOS LIBERADOS SCS
echo ============================================================
echo.
echo [1] Version de PowerShell:
powershell.exe -NoProfile -Command "$PSVersionTable.PSVersion"
echo.
echo [2] Acceso a carpeta:
powershell.exe -NoProfile -Command "$p='\\gerencialz.canariasalud\archivos\Logistica\COMPRAS\Pedidos Liberados'; Write-Host $p; if(Test-Path -LiteralPath $p){Write-Host 'OK' -ForegroundColor Green}else{Write-Host 'SIN ACCESO' -ForegroundColor Red}"
echo.
echo [3] Tarea programada:
schtasks /Query /TN "SCS - Pedidos Liberados" /V /FO LIST 2>nul
echo.
echo [4] Configuracion de correo:
powershell.exe -NoProfile -Command "$c=Get-Content -Raw '%~dp0config.json'|ConvertFrom-Json; $c.Mail | Format-List Enabled,SmtpHost,SmtpPort,EnableSsl,AuthenticationMode,FromAddress"
echo.
pause
