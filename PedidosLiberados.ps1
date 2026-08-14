param(
    [switch]$Manual
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptRoot "config.json"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    try {
        $stateFolder = $script:StateFolder
        if (-not $stateFolder) { return }
        New-Item -ItemType Directory -Path $stateFolder -Force | Out-Null
        $logPath = Join-Path $stateFolder ("PedidosLiberados_{0}.log" -f (Get-Date -Format "yyyyMMdd"))
        $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
        Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    } catch {}
}

function Expand-EnvPath([string]$Path) {
    return [Environment]::ExpandEnvironmentVariables($Path)
}

function Replace-File {
    param([string]$Source, [string]$Destination)
    if (Test-Path -LiteralPath $Destination) {
        $backup = "$Destination.$([guid]::NewGuid().ToString('N')).bak"
        [IO.File]::Replace($Source, $Destination, $backup)
        Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue
    } else {
        Move-Item -LiteralPath $Source -Destination $Destination
    }
}

function Get-State {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $obj = $raw | ConvertFrom-Json
        if ($null -eq $obj) { return @() }
        if (@($obj).Count -eq 1 -and
            $obj[0].PSObject.Properties.Name -contains "value" -and
            $obj[0].PSObject.Properties.Name -contains "Count") {
            return @($obj[0].value)
        }
        return @($obj)
    } catch {
        Write-Log "No se pudo leer el estado: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Save-State {
    param([object[]]$State, [string]$Path)

    $tmp = "$Path.tmp"
    $json = ConvertTo-Json -InputObject ([object[]]@($State)) -Depth 8
    Set-Content -LiteralPath $tmp -Value $json -Encoding UTF8
    Replace-File $tmp $Path
}

function Find-StateRecord {
    param([object[]]$State, [string]$FileName)
    return @($State | Where-Object { $_.FileName -eq $FileName } | Select-Object -First 1)[0]
}

function New-StateRecord {
    param([System.IO.FileInfo]$File, [hashtable]$Parsed)
    return [pscustomobject][ordered]@{
        FileName       = $File.Name
        FullPath       = $File.FullName
        Hash           = ""
        Usuario        = $Parsed.Usuario
        PRD            = $Parsed.PRD
        Impresion      = [int]$Parsed.Impresion
        Proveedor      = $Parsed.Proveedor
        Pedido         = $Parsed.Pedido
        Empresa        = ""
        Para           = ""
        CC             = ""
        Estado         = "PENDIENTE"
        Intentos       = 0
        NextAttempt    = ""
        Error          = ""
        CreatedAt      = (Get-Date).ToString("o")
        UpdatedAt      = (Get-Date).ToString("o")
        SentAt         = ""
        ArchivedAt     = ""
    }
}

function Add-Event {
    param([object]$Record, [string]$Event, [string]$Detail = "")
    try {
        $folder = Join-Path $script:StateFolder "eventos"
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $path = Join-Path $folder ("eventos_{0}.jsonl" -f (Get-Date -Format "yyyyMMdd"))
        $e = [pscustomobject][ordered]@{
            FechaHora = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Evento = $Event
            Pedido = $Record.Pedido
            Proveedor = $Record.Proveedor
            Empresa = $Record.Empresa
            Usuario = $Record.Usuario
            PRD = $Record.PRD
            Impresion = $Record.Impresion
            Archivo = $Record.FileName
            Estado = $Record.Estado
            Intentos = $Record.Intentos
            Error = if ($Detail) { $Detail } else { $Record.Error }
            Para = $Record.Para
            CC = $Record.CC
            Hash = $Record.Hash
        }
        Add-Content -LiteralPath $path -Value ($e | ConvertTo-Json -Compress -Depth 5) -Encoding UTF8
    } catch {
        Write-Log "No se pudo registrar evento: $($_.Exception.Message)" "WARN"
    }
}

function Parse-OrderFileName {
    param([string]$Name)
    $rx = '^(?<usr>[^_]+)_(?<prd>PRD\d+)_(?<imp>\d+)_Proveedor_\s*(?<prov>\d+)\s+N.\s+Pedido_\s*(?<ped>\d{10})\.pdf$'
    $m = [regex]::Match($Name, $rx, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $m.Success) { return $null }
    return @{
        Usuario   = $m.Groups["usr"].Value.Trim()
        PRD       = $m.Groups["prd"].Value.Trim()
        Impresion = [int]$m.Groups["imp"].Value
        Proveedor = $m.Groups["prov"].Value.Trim()
        Pedido    = $m.Groups["ped"].Value.Trim()
    }
}

function Test-FileReady {
    param([System.IO.FileInfo]$File, [int]$StableSeconds)
    if (((Get-Date) - $File.LastWriteTime).TotalSeconds -lt $StableSeconds) { return $false }
    try {
        $s = [IO.File]::Open($File.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $s.Close()
        return $true
    } catch { return $false }
}

function Get-FileSha256 {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Get-ColIndex([string]$CellRef) {
    $letters = ([regex]::Match($CellRef, '^[A-Z]+')).Value
    $n = 0
    foreach ($c in $letters.ToCharArray()) { $n = ($n * 26) + ([int][char]$c - [int][char]'A' + 1) }
    return $n
}

function Read-ZipText {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $reader = New-Object IO.StreamReader($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Import-XlsxFirstSheet {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $fileStream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read,
        [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    $zip = New-Object IO.Compression.ZipArchive($fileStream, [IO.Compression.ZipArchiveMode]::Read, $false)
    try {
        $shared = @()
        $sstText = Read-ZipText $zip "xl/sharedStrings.xml"
        if ($sstText) {
            [xml]$sst = $sstText
            foreach ($si in $sst.sst.si) {
                $shared += [string]$si.InnerText
            }
        }

        [xml]$wb = Read-ZipText $zip "xl/workbook.xml"
        [xml]$rels = Read-ZipText $zip "xl/_rels/workbook.xml.rels"
        $sheet = @($wb.workbook.sheets.sheet)[0]
        $rid = $sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $rel = @($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid })[0]
        if ($null -eq $rel) { throw "No se pudo localizar la primera hoja de Excel." }
        $target = [string]$rel.Target
        if ($target.StartsWith("/")) { $sheetPath = $target.TrimStart("/") }
        elseif ($target.StartsWith("xl/")) { $sheetPath = $target }
        else { $sheetPath = "xl/" + $target.TrimStart("/") }

        [xml]$ws = Read-ZipText $zip $sheetPath
        $matrix = @()
        foreach ($row in @($ws.SelectNodes("/*[local-name()='worksheet']/*[local-name()='sheetData']/*[local-name()='row']"))) {
            $values = @{}
            $maxCol = 0
            foreach ($cell in @($row.SelectNodes("./*[local-name()='c']"))) {
                $idx = Get-ColIndex ([string]$cell.GetAttribute("r"))
                if ($idx -gt $maxCol) { $maxCol = $idx }
                $t = [string]$cell.GetAttribute("t")
                $val = ""
                if ($t -eq "s") {
                    $valueNode = $cell.SelectSingleNode("./*[local-name()='v']")
                    if ($valueNode) {
                        $i = [int]$valueNode.InnerText
                        if ($i -ge 0 -and $i -lt $shared.Count) { $val = $shared[$i] }
                    }
                } elseif ($t -eq "inlineStr") {
                    $inline = $cell.SelectSingleNode("./*[local-name()='is']")
                    if ($inline) {
                        $val = [string]$inline.InnerText
                    }
                } else {
                    $valueNode = $cell.SelectSingleNode("./*[local-name()='v']")
                    if ($valueNode) { $val = [string]$valueNode.InnerText }
                }
                $values[$idx] = $val
            }
            if ($maxCol -gt 0) {
                $arr = New-Object string[] $maxCol
                foreach ($k in $values.Keys) { $arr[$k-1] = [string]$values[$k] }
                $matrix += ,$arr
            }
        }
        if ($matrix.Count -lt 1) { return @() }
        $headers = @($matrix[0] | ForEach-Object { ([string]$_).Trim() })
        $result = @()
        for ($r=1; $r -lt $matrix.Count; $r++) {
            $o = [ordered]@{}
            for ($c=0; $c -lt $headers.Count; $c++) {
                if ([string]::IsNullOrWhiteSpace($headers[$c])) { continue }
                $v = if ($c -lt $matrix[$r].Count) { [string]$matrix[$r][$c] } else { "" }
                $o[$headers[$c]] = $v.Trim()
            }
            if ($o.Count -gt 0) { $result += [pscustomobject]$o }
        }
        return $result
    } finally {
        $zip.Dispose()
        $fileStream.Dispose()
    }
}

function Split-MailAddresses {
    param([string]$Value)
    $result = @()
    foreach ($item in @([string]$Value -split '[;,]')) {
        $address = $item.Trim()
        if (-not $address) { continue }
        try {
            $parsed = New-Object System.Net.Mail.MailAddress($address)
            $result += $parsed.Address
        } catch {
            throw "Direccion de correo no valida: $address"
        }
    }
    return @($result)
}

function Get-Provider {
    param([string]$ProvidersPath, [string]$Code)
    $rows = Import-XlsxFirstSheet $ProvidersPath
    foreach ($r in $rows) {
        if (-not ($r.PSObject.Properties.Name -contains "COD_PROVEEDOR")) { continue }
        if (([string]$r.COD_PROVEEDOR).Trim() -ne $Code.Trim()) { continue }
        $activo = if ($r.PSObject.Properties.Name -contains "ACTIVO") { ([string]$r.ACTIVO).Trim().ToUpperInvariant() } else { "SI" }
        if ($activo -notin @("SI","SÍ","YES","TRUE","1","")) { return $null }

        $primary = @()
        if ($r.PSObject.Properties.Name -contains "EMAIL_1") { $primary = @(Split-MailAddresses ([string]$r.EMAIL_1)) }
        $cc = @()
        $emailProps = @($r.PSObject.Properties.Name | Where-Object { $_ -match '^EMAIL_(\d+)$' } |
            Sort-Object { [int]([regex]::Match($_,'\d+').Value) })
        foreach ($p in $emailProps) {
            if ($p -eq "EMAIL_1") { continue }
            $cc += @(Split-MailAddresses ([string]$r.$p))
        }
        return [pscustomobject]@{
            Code = $Code
            Name = if ($r.PSObject.Properties.Name -contains "NOMBRE_PROVEEDOR") { ([string]$r.NOMBRE_PROVEEDOR).Trim() } else { "" }
            Primary = @($primary)
            CC = @($cc)
        }
    }
    return $null
}

function Send-OrderMail {
    param([object]$Record, [string]$PdfPath, $MailCfg)

    if (-not [bool]$MailCfg.Enabled) { throw "El envio de correo esta deshabilitado en config.json (Mail.Enabled=false)." }
    if ([string]$MailCfg.SmtpHost -eq "PENDIENTE_CONFIGURAR") { throw "Servidor SMTP pendiente de configurar." }

    $msg = New-Object System.Net.Mail.MailMessage
    $smtp = $null
    try {
        $msg.From = New-Object System.Net.Mail.MailAddress([string]$MailCfg.FromAddress, [string]$MailCfg.FromName)
        foreach ($addr in (Split-MailAddresses ([string]$Record.Para))) {
            $msg.To.Add($addr)
        }
        if ($msg.To.Count -eq 0) { throw "El proveedor no tiene destinatarios principales validos." }
        if ($Record.CC) {
            foreach ($addr in (Split-MailAddresses ([string]$Record.CC))) {
                $msg.CC.Add($addr)
            }
        }

        $orgName = if ($MailCfg.PSObject.Properties.Name -contains "OrganizationName") { [string]$MailCfg.OrganizationName } else { "" }
        $msg.Subject = ("Pedido " + $orgName).TrimEnd() + " n" + [char]186 + " " + $Record.Pedido + " " + [char]8211 + " " + $Record.Empresa

        $signature = ""
        $sigPath = Join-Path $ScriptRoot ([string]$MailCfg.SignatureFile)
        if (Test-Path -LiteralPath $sigPath) {
            $signature = Get-Content -LiteralPath $sigPath -Raw -Encoding UTF8
        }

        $empresaEsc = [System.Net.WebUtility]::HtmlEncode([string]$Record.Empresa)
        $pedidoEsc = [System.Net.WebUtility]::HtmlEncode([string]$Record.Pedido)
        $msg.Body = "<p>Adjunto remitimos Pedido n&ordm;: " + $pedidoEsc + " para " + $empresaEsc + "</p>" +
                    "<p>Rogamos confirmaci&oacute;n de la recepci&oacute;n de este correo y/o activar en su cuenta de correo electr&oacute;nico la confirmaci&oacute;n autom&aacute;tica de recepci&oacute;n de correos.</p>" +
                    "<p>En caso de incidencias con el pedido contactar con: " + [System.Net.WebUtility]::HtmlEncode([string]$MailCfg.FromAddress) + "</p>" +
                    "<p>Atentamente,</p>" + $signature
        $msg.IsBodyHtml = $true
        $msg.BodyEncoding = [Text.Encoding]::UTF8
        $msg.SubjectEncoding = [Text.Encoding]::UTF8

        if ([bool]$MailCfg.RequestReadReceipt) {
            $msg.Headers.Add("Disposition-Notification-To", [string]$MailCfg.FromAddress)
            $msg.Headers.Add("Return-Receipt-To", [string]$MailCfg.FromAddress)
        }
        if ([bool]$MailCfg.RequestDeliveryReceipt) {
            $msg.DeliveryNotificationOptions = [System.Net.Mail.DeliveryNotificationOptions]::OnFailure -bor `
                                               [System.Net.Mail.DeliveryNotificationOptions]::Delay
        }

        $att = New-Object System.Net.Mail.Attachment($PdfPath)
        $msg.Attachments.Add($att)

        $smtp = New-Object System.Net.Mail.SmtpClient([string]$MailCfg.SmtpHost, [int]$MailCfg.SmtpPort)
        $smtp.EnableSsl = [bool]$MailCfg.EnableSsl
        $smtp.Timeout = 120000

        if (([string]$MailCfg.AuthenticationMode) -eq "WindowsIntegrated") {
            $smtp.UseDefaultCredentials = $true
        } else {
            $user = [string]$MailCfg.Username
            $envName = [string]$MailCfg.PasswordEnvironmentVariable
            $password = [Environment]::GetEnvironmentVariable($envName, "Machine")
            if (-not $password) { $password = [Environment]::GetEnvironmentVariable($envName, "Process") }
            if (-not $password) { throw "No existe la variable de entorno $envName con la contrasena SMTP." }
            $smtp.UseDefaultCredentials = $false
            $smtp.Credentials = New-Object System.Net.NetworkCredential($user, $password)
        }
        $smtp.Send($msg)
    } finally {
        if ($smtp) { $smtp.Dispose() }
        $msg.Dispose()
    }
}

function Xml-Escape([string]$s) {
    if ($null -eq $s) { return "" }
    return [Security.SecurityElement]::Escape([string]$s)
}
function Col-Letter([int]$n) {
    $s = ""
    while ($n -gt 0) {
        $n--
        $s = [char](65 + ($n % 26)) + $s
        $n = [math]::Floor($n / 26)
    }
    return $s
}
function Export-SimpleXlsx {
    param([string]$Path, [object[]]$Rows, [string[]]$Columns)
    try {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $tmpRoot = Join-Path $env:TEMP ("PedidosXlsx_" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path (Join-Path $tmpRoot "_rels") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tmpRoot "xl\_rels") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tmpRoot "xl\worksheets") -Force | Out-Null

        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@ | Set-Content -LiteralPath (Join-Path $tmpRoot "[Content_Types].xml") -Encoding UTF8

        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@ | Set-Content -LiteralPath (Join-Path $tmpRoot "_rels\.rels") -Encoding UTF8

        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Datos" sheetId="1" r:id="rId1"/></sheets>
</workbook>
'@ | Set-Content -LiteralPath (Join-Path $tmpRoot "xl\workbook.xml") -Encoding UTF8

        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@ | Set-Content -LiteralPath (Join-Path $tmpRoot "xl\_rels\workbook.xml.rels") -Encoding UTF8

        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
</styleSheet>
'@ | Set-Content -LiteralPath (Join-Path $tmpRoot "xl\styles.xml") -Encoding UTF8

        $sb = New-Object Text.StringBuilder
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        [void]$sb.Append('<cols>')
        for ($i=1; $i -le $Columns.Count; $i++) {
            $w = if ($Columns[$i-1] -in @("Archivo","Error")) { 55 } elseif ($Columns[$i-1] -in @("Empresa","Para","CC")) { 35 } else { 18 }
            [void]$sb.Append("<col min=`"$i`" max=`"$i`" width=`"$w`" customWidth=`"1`"/>")
        }
        [void]$sb.Append('</cols><sheetData>')
        [void]$sb.Append('<row r="1">')
        for ($c=0; $c -lt $Columns.Count; $c++) {
            $ref = (Col-Letter ($c+1)) + "1"
            [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`" s=`"1`"><is><t>$(Xml-Escape $Columns[$c])</t></is></c>")
        }
        [void]$sb.Append('</row>')
        $rnum = 2
        foreach ($row in @($Rows)) {
            [void]$sb.Append("<row r=`"$rnum`">")
            for ($c=0; $c -lt $Columns.Count; $c++) {
                $name = $Columns[$c]
                $v = ""
                if ($row -and ($row.PSObject.Properties.Name -contains $name)) { $v = [string]$row.$name }
                $ref = (Col-Letter ($c+1)) + $rnum
                [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(Xml-Escape $v)</t></is></c>")
            }
            [void]$sb.Append('</row>')
            $rnum++
        }
        [void]$sb.Append('</sheetData><autoFilter ref="A1:' + (Col-Letter $Columns.Count) + '1"/></worksheet>')
        $sb.ToString() | Set-Content -LiteralPath (Join-Path $tmpRoot "xl\worksheets\sheet1.xml") -Encoding UTF8

        $tempZip = "$Path.tmp"
        if (Test-Path -LiteralPath $tempZip) { Remove-Item -LiteralPath $tempZip -Force }
        $archive = [IO.Compression.ZipFile]::Open($tempZip, [IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($sourceFile in Get-ChildItem -LiteralPath $tmpRoot -Recurse -File) {
                $entryName = $sourceFile.FullName.Substring($tmpRoot.Length).TrimStart('\','/') -replace '\\','/'
                [IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                    $archive, $sourceFile.FullName, $entryName, [IO.Compression.CompressionLevel]::Optimal
                ) | Out-Null
            }
        } finally {
            $archive.Dispose()
        }
        Replace-File $tempZip $Path
    } finally {
        if ($tmpRoot -and (Test-Path -LiteralPath $tmpRoot)) { Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Update-Reports {
    param([object[]]$State, [string]$ReportsFolder)
    New-Item -ItemType Directory -Path $ReportsFolder -Force | Out-Null

    $cols = @("FechaHora","Evento","Pedido","Proveedor","Empresa","Usuario","PRD","Impresion","Archivo","Estado","Intentos","Error","Para","CC","Hash")
    $eventFile = Join-Path (Join-Path $script:StateFolder "eventos") ("eventos_{0}.jsonl" -f (Get-Date -Format "yyyyMMdd"))
    $events = @()
    if (Test-Path -LiteralPath $eventFile) {
        foreach ($line in Get-Content -LiteralPath $eventFile -Encoding UTF8) {
            try { $events += ($line | ConvertFrom-Json) } catch {}
        }
    }
    $daily = Join-Path $ReportsFolder ("REPORT_PEDIDOS_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd"))
    try { Export-SimpleXlsx $daily $events $cols } catch { Write-Log ("No se pudo actualizar " + $daily + ": " + $_.Exception.Message) "WARN" }

    $incidentStates = @("INCIDENCIA_PERMANENTE","INCIDENCIA_NOMBRE","REVISION_MANUAL","ERROR_ARCHIVO","DUPLICADO","ESPERA_DATOS")
    $inc = @()
    foreach ($r in @($State | Where-Object { $_.Estado -in $incidentStates })) {
        $inc += [pscustomobject][ordered]@{
            FechaHora = $r.UpdatedAt
            Evento = "INCIDENCIA_ACTIVA"
            Pedido = $r.Pedido
            Proveedor = $r.Proveedor
            Empresa = $r.Empresa
            Usuario = $r.Usuario
            PRD = $r.PRD
            Impresion = $r.Impresion
            Archivo = $r.FileName
            Estado = $r.Estado
            Intentos = $r.Intentos
            Error = $r.Error
            Para = $r.Para
            CC = $r.CC
            Hash = $r.Hash
        }
    }
    $incPath = Join-Path $ReportsFolder "INCIDENCIAS_PEDIDOS.xlsx"
    try { Export-SimpleXlsx $incPath $inc $cols } catch { Write-Log "No se pudo actualizar incidencias: $($_.Exception.Message)" "WARN" }
}

function Update-Dashboard {
    param([object[]]$State, [string]$ReportsFolder)
    try {
        $pending = @($State | Where-Object { $_.Estado -in @("PENDIENTE","ESPERA_DATOS","ERROR_TEMPORAL","ENVIADO_PENDIENTE_ARCHIVO") }).Count
        $incidents = @($State | Where-Object { $_.Estado -in @("INCIDENCIA_PERMANENTE","INCIDENCIA_NOMBRE","REVISION_MANUAL","ERROR_ARCHIVO","DUPLICADO","ESPERA_DATOS") }).Count
        $today = (Get-Date).ToString("yyyy-MM-dd")
        $sent = @($State | Where-Object { $_.SentAt -and ([string]$_.SentAt).StartsWith($today) }).Count

        $rows = New-Object Text.StringBuilder
        foreach ($r in @($State | Sort-Object UpdatedAt -Descending | Select-Object -First 100)) {
            [void]$rows.Append("<tr><td>$(Xml-Escape $r.Pedido)</td><td>$(Xml-Escape $r.Proveedor)</td><td>$(Xml-Escape $r.Empresa)</td><td>$(Xml-Escape $r.Estado)</td><td>$($r.Intentos)</td><td>$(Xml-Escape $r.Error)</td><td>$(Xml-Escape $r.FileName)</td></tr>")
        }
        $html = @"
<!doctype html><html><head><meta charset="utf-8"><title>Pedidos Liberados SCS</title>
<style>body{font-family:Segoe UI,Arial;margin:28px;color:#1f2937}h1{margin-bottom:8px}.kpi{display:inline-block;padding:16px 22px;margin:8px 10px 18px 0;border:1px solid #d1d5db;border-radius:8px}.n{font-size:28px;font-weight:700}table{border-collapse:collapse;width:100%;font-size:13px}th,td{border-bottom:1px solid #ddd;padding:8px;text-align:left;vertical-align:top}th{background:#f3f4f6;position:sticky;top:0}.small{color:#6b7280}</style>
</head><body><h1>Envio automatico de Pedidos Liberados</h1>
<div class="small">Ultima ejecucion: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") - Ejecucion programada cada 5 minutos</div>
<div class="kpi"><div class="n">$pending</div>Pendientes</div>
<div class="kpi"><div class="n">$sent</div>Enviados hoy</div>
<div class="kpi"><div class="n">$incidents</div>Incidencias</div>
<p><b>Para forzar una revision inmediata:</b> ejecute <code>EJECUTAR_AHORA.cmd</code>.</p>
<table><thead><tr><th>Pedido</th><th>Proveedor</th><th>Empresa</th><th>Estado</th><th>Intentos</th><th>Error</th><th>Archivo</th></tr></thead><tbody>$($rows.ToString())</tbody></table>
</body></html>
"@
        $panel = Join-Path $ReportsFolder "panel_control.html"
        $html | Set-Content -LiteralPath $panel -Encoding UTF8
    } catch { Write-Log "No se pudo actualizar panel: $($_.Exception.Message)" "WARN" }
}

# -------------------- MAIN --------------------
if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "No se encuentra config.json." }
$cfg = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$root = [string]$cfg.Pedidos.RootPath
$script:StateFolder = Expand-EnvPath ([string]$cfg.Pedidos.StateFolder)
New-Item -ItemType Directory -Path $script:StateFolder -Force | Out-Null
$statePath = Join-Path $script:StateFolder "estado.json"

try {
    if (-not (Test-Path -LiteralPath $root)) { throw "No se puede acceder a la carpeta raiz: $root" }

    $providersPath = Join-Path $root ([string]$cfg.Pedidos.ProvidersFile)
    $sentRoot = Join-Path $root ([string]$cfg.Pedidos.SentFolder)
    $reprintRoot = Join-Path $root ([string]$cfg.Pedidos.ReprintsFolder)
    $reportsRoot = Join-Path $root ([string]$cfg.Pedidos.ReportsFolder)
    New-Item -ItemType Directory -Path $sentRoot,$reprintRoot,$reportsRoot -Force | Out-Null

    # If local package has Proveedores.xlsx and network copy does not yet exist, seed it once.
    if (-not (Test-Path -LiteralPath $providersPath)) {
        $localProviders = Join-Path $ScriptRoot ([string]$cfg.Pedidos.ProvidersFile)
        if (Test-Path -LiteralPath $localProviders) {
            Copy-Item -LiteralPath $localProviders -Destination $providersPath
            Write-Log "Se creo la base inicial de proveedores en $providersPath"
        }
    }

    $state = @(Get-State $statePath)

    # Safety after an interrupted send: never retry automatically an indeterminate ENVIANDO.
    foreach ($r in @($state | Where-Object { $_.Estado -eq "ENVIANDO" })) {
        $r.Estado = "REVISION_MANUAL"
        $r.Error = "La ejecucion anterior termino mientras el pedido estaba ENVIANDO. Revisar manualmente si Exchange lo acepto antes de cualquier reenvio."
        $r.UpdatedAt = (Get-Date).ToString("o")
        Add-Event $r "REVISION_MANUAL" $r.Error
    }
    Save-State $state $statePath

    # Finish archive-only operations first.
    foreach ($r in @($state | Where-Object { $_.Estado -eq "ENVIADO_PENDIENTE_ARCHIVO" })) {
        if (-not (Test-Path -LiteralPath $r.FullPath)) { continue }
        try {
            $destDir = Join-Path (Join-Path $sentRoot (Get-Date -Format "yyyy")) (Get-Date -Format "MM")
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            $dest = Join-Path $destDir $r.FileName
            if (Test-Path -LiteralPath $dest) { throw "Ya existe un archivo con el mismo nombre en destino." }
            Move-Item -LiteralPath $r.FullPath -Destination $dest
            $r.FullPath = $dest
            $r.Estado = "ARCHIVADO"; $r.ArchivedAt = (Get-Date).ToString("o"); $r.UpdatedAt = (Get-Date).ToString("o"); $r.Error = ""
            Add-Event $r "ARCHIVADO"
            Save-State $state $statePath
        } catch {
            $r.Error = "Correo enviado, pero no se pudo archivar: $($_.Exception.Message)"
            $r.UpdatedAt = (Get-Date).ToString("o")
            Add-Event $r "ERROR_ARCHIVO" $r.Error
            Save-State $state $statePath
        }
    }

    $files = @(Get-ChildItem -LiteralPath $root -Filter "*.pdf" -File -ErrorAction Stop)
    foreach ($file in $files) {
        if (-not (Test-FileReady $file ([int]$cfg.Pedidos.StableSeconds))) { continue }

        $parsed = Parse-OrderFileName $file.Name
        $record = Find-StateRecord $state $file.Name

        if ($null -eq $parsed) {
            if ($null -eq $record) {
                $record = [pscustomobject][ordered]@{
                    FileName=$file.Name; FullPath=$file.FullName; Hash=""; Usuario=""; PRD=""; Impresion=0; Proveedor=""; Pedido="";
                    Empresa=""; Para=""; CC=""; Estado="INCIDENCIA_NOMBRE"; Intentos=0; NextAttempt="";
                    Error="El nombre del PDF no cumple el patron esperado."; CreatedAt=(Get-Date).ToString("o"); UpdatedAt=(Get-Date).ToString("o"); SentAt=""; ArchivedAt=""
                }
                $state += $record
                Add-Event $record "INCIDENCIA_NOMBRE" $record.Error
                Save-State $state $statePath
            }
            continue
        }

        if ([int]$parsed.Impresion -ne 1) {
            try {
                $dest = Join-Path $reprintRoot $file.Name
                if (Test-Path -LiteralPath $dest) { throw "Ya existe la reimpresion en destino." }
                Move-Item -LiteralPath $file.FullName -Destination $dest
                if ($null -eq $record) {
                    $record = New-StateRecord $file $parsed
                    $state += $record
                }
                $record.FullPath=$dest
                $record.Estado="REIMPRESION"; $record.UpdatedAt=(Get-Date).ToString("o"); $record.Error=""
                Add-Event $record "REIMPRESION"
                Save-State $state $statePath
            } catch {
                if ($null -eq $record) { $record = New-StateRecord $file $parsed; $state += $record }
                $record.Estado="ERROR_ARCHIVO"; $record.Error=$_.Exception.Message; $record.UpdatedAt=(Get-Date).ToString("o")
                Add-Event $record "ERROR_REIMPRESION" $record.Error
                Save-State $state $statePath
            }
            continue
        }

        if ($null -eq $record) {
            $record = New-StateRecord $file $parsed
            $state += $record
        } else {
            $record.FullPath = $file.FullName
        }

        if ($record.Estado -in @("ARCHIVADO","REIMPRESION","REVISION_MANUAL","INCIDENCIA_PERMANENTE","DUPLICADO")) { continue }
        if ($record.NextAttempt) {
            try {
                if ((Get-Date) -lt [datetime]$record.NextAttempt) { continue }
            } catch {}
        }

        try {
            $record.Hash = Get-FileSha256 $file.FullName

            # Extra duplicate guard: same hash already archived/sent with another filename.
            $dup = @($state | Where-Object {
                $_ -ne $record -and $_.Hash -eq $record.Hash -and $_.Estado -in @("ARCHIVADO","ENVIADO_PENDIENTE_ARCHIVO")
            } | Select-Object -First 1)
            if ($dup.Count -gt 0) {
                $record.Estado="DUPLICADO"; $record.Error="El mismo PDF ya consta como enviado/archivado: $($dup[0].FileName)"; $record.UpdatedAt=(Get-Date).ToString("o")
                Add-Event $record "DUPLICADO" $record.Error
                Save-State $state $statePath
                continue
            }

            if (-not (Test-Path -LiteralPath $providersPath)) { throw "No existe Proveedores.xlsx en la carpeta raiz." }
            $provider = Get-Provider $providersPath $record.Proveedor
            if ($null -eq $provider) {
                $record.Estado="ESPERA_DATOS"; $record.Error="Proveedor no encontrado o inactivo en Proveedores.xlsx."; $record.UpdatedAt=(Get-Date).ToString("o"); $record.NextAttempt=""
                Add-Event $record "ESPERA_DATOS" $record.Error
                Save-State $state $statePath
                continue
            }
            if (@($provider.Primary).Count -eq 0) {
                $record.Estado="ESPERA_DATOS"; $record.Error="Proveedor localizado, pero EMAIL_1 no esta informado."; $record.UpdatedAt=(Get-Date).ToString("o"); $record.NextAttempt=""
                $record.Empresa=$provider.Name
                Add-Event $record "ESPERA_DATOS" $record.Error
                Save-State $state $statePath
                continue
            }

            $record.Empresa = $provider.Name
            $record.Para = (@($provider.Primary) -join ";")
            $record.CC = ($provider.CC -join ";")
            $record.Estado = "ENVIANDO"
            $record.Error = ""
            $record.UpdatedAt = (Get-Date).ToString("o")
            Add-Event $record "ENVIANDO"
            Save-State $state $statePath

            try {
                Send-OrderMail $record $file.FullName $cfg.Mail
                $record.Estado = "ENVIADO_PENDIENTE_ARCHIVO"
                $record.SentAt = (Get-Date).ToString("o")
                $record.UpdatedAt = (Get-Date).ToString("o")
                $record.NextAttempt = ""
                $record.Error = ""
                Add-Event $record "ENVIADO"
                Save-State $state $statePath

                $destDir = Join-Path (Join-Path $sentRoot (Get-Date -Format "yyyy")) (Get-Date -Format "MM")
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                $dest = Join-Path $destDir $file.Name
                if (Test-Path -LiteralPath $dest) { throw "Ya existe un archivo con el mismo nombre en pedidos enviados." }
                Move-Item -LiteralPath $file.FullName -Destination $dest
                $record.FullPath=$dest
                $record.Estado="ARCHIVADO"; $record.ArchivedAt=(Get-Date).ToString("o"); $record.UpdatedAt=(Get-Date).ToString("o")
                Add-Event $record "ARCHIVADO"
                Save-State $state $statePath
            } catch {
                # If mail was accepted, state has already become ENVIADO_PENDIENTE_ARCHIVO.
                if ($record.Estado -eq "ENVIADO_PENDIENTE_ARCHIVO") {
                    $record.Error = "Correo enviado, pero no se pudo archivar: $($_.Exception.Message)"
                    $record.UpdatedAt=(Get-Date).ToString("o")
                    Add-Event $record "ERROR_ARCHIVO" $record.Error
                    Save-State $state $statePath
                } else {
                    $record.Intentos = [int]$record.Intentos + 1
                    $record.Error = $_.Exception.Message
                    $record.UpdatedAt = (Get-Date).ToString("o")
                    if ($record.Intentos -ge [int]$cfg.Pedidos.MaxMailAttempts) {
                        $record.Estado = "INCIDENCIA_PERMANENTE"
                        $record.NextAttempt = ""
                    } else {
                        $record.Estado = "ERROR_TEMPORAL"
                        $mins = if ($record.Intentos -eq 1) { [int]$cfg.Pedidos.Retry1Minutes } else { [int]$cfg.Pedidos.Retry2Minutes }
                        $record.NextAttempt = (Get-Date).AddMinutes($mins).ToString("o")
                    }
                    Add-Event $record "ERROR_ENVIO" $record.Error
                    Save-State $state $statePath
                }
            }
        } catch {
            $record.Estado="ESPERA_DATOS"; $record.Error=$_.Exception.Message; $record.UpdatedAt=(Get-Date).ToString("o")
            Add-Event $record "ERROR_PREVALIDACION" $record.Error
            Save-State $state $statePath
        }
    }

    Update-Reports $state $reportsRoot
    Update-Dashboard $state $reportsRoot
    Write-Log "Ciclo completado. PDFs encontrados: $($files.Count)"
} catch {
    Write-Log $_.Exception.ToString() "FATAL"
    throw
}
