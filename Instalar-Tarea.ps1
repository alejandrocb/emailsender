$ErrorActionPreference = "Stop"
$taskName = "SCS - Pedidos Liberados"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $root "PedidosLiberados.ps1"

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Ejecute INSTALAR_TAREA.cmd como administrador."
}

Write-Host ""
Write-Host "INSTALACION - PEDIDOS LIBERADOS SCS" -ForegroundColor Cyan
Write-Host "La tarea se ejecutara cada 5 minutos con una cuenta corporativa." -ForegroundColor Gray
Write-Host ""

$defaultUser = "$env:USERDOMAIN\$env:USERNAME"
$user = Read-Host "Cuenta de servicio [ej. DOMINIO\usuario] (Enter = $defaultUser)"
if ([string]::IsNullOrWhiteSpace($user)) { $user = $defaultUser }

$cred = Get-Credential -UserName $user -Message "Introduzca la contrasena de la cuenta que ejecutara la tarea. Windows Task Scheduler la almacenara de forma protegida."

$arg = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$script`""
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $arg -WorkingDirectory $root
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 5)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 4)

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
    -User $cred.UserName -Password $cred.GetNetworkCredential().Password -RunLevel Highest | Out-Null

Write-Host ""
Write-Host "[OK] Tarea instalada: $taskName" -ForegroundColor Green
Write-Host "Se ejecutara cada 5 minutos, aunque no haya una sesion interactiva abierta." -ForegroundColor Green
Write-Host ""
Write-Host "IMPORTANTE: Mail.Enabled permanece en FALSE hasta configurar Exchange y realizar una prueba controlada." -ForegroundColor Yellow
Write-Host "Abra CONFIGURAR_CORREO.cmd para editar config.json." -ForegroundColor Yellow
Write-Host ""
Write-Host "Puede probar lectura, clasificacion e informes sin enviar correos ejecutando EJECUTAR_AHORA.cmd." -ForegroundColor Cyan
