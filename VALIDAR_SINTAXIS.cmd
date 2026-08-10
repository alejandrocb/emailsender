@echo off
chcp 65001 >nul
echo ============================================================
echo   VALIDACION DE SINTAXIS - PEDIDOS LIBERADOS SCS
echo ============================================================
echo.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile('%~dp0PedidosLiberados.ps1',[ref]$tokens,[ref]$errors) ^> $null; if($errors.Count -eq 0){Write-Host '[OK] Sintaxis PowerShell valida.' -ForegroundColor Green; exit 0}else{$errors ^| Format-Table -AutoSize Message,Extent; exit 1}"
set RC=%ERRORLEVEL%
echo.
if "%RC%"=="0" (
  echo Puede ejecutar EJECUTAR_AHORA.cmd.
) else (
  echo NO ejecute el programa. Remita el resultado del error.
)
echo.
pause
exit /b %RC%
