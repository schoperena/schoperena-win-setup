<#
.SYNOPSIS
    Genera codigos QR (enlaces, WiFi, contacto, email, SMS, geo...) con colores personalizables.

.DESCRIPTION
    Sin argumentos abre un menu interactivo. Con -Text genera directamente.
    Todo se genera localmente (libreria QRCoder); ningun dato sale del equipo.

.EXAMPLE
    .\New-QRCode.ps1
    Menu interactivo.

.EXAMPLE
    .\New-QRCode.ps1 -Text "https://github.com/schoperena" -Dark "#0D1117" -Light "#58A6FF"
    Genera un QR con colores personalizados sin preguntar nada.

.EXAMPLE
    .\New-QRCode.ps1 -Text "hola" -Formato svg -Out .\hola.svg
#>

[CmdletBinding()]
param(
    # Contenido a codificar. Si se omite, se abre el menu interactivo.
    [string] $Text,

    # Ruta del archivo de salida. Por defecto: Escritorio\qr-<slug>-<fecha>.<ext>
    [string] $Out,

    # Color de los modulos (nombre, #RGB o #RRGGBB).
    [string] $Dark = 'negro',

    # Color del fondo (nombre, #RGB, #RRGGBB o "transparente").
    [string] $Light = 'blanco',

    # Pixeles por modulo. Mas alto = imagen mas grande.
    [ValidateRange(1, 100)]
    [int] $Escala = 20,

    # Nivel de correccion de errores: L(7%) M(15%) Q(25%) H(30%).
    [ValidateSet('L', 'M', 'Q', 'H')]
    [string] $Ecc = 'Q',

    [ValidateSet('png', 'svg')]
    [string] $Formato = 'png',

    # Quita el margen blanco alrededor del QR (no recomendado: dificulta el escaneo).
    [switch] $SinMargen,

    # No dibuja la vista previa en la consola.
    [switch] $SinPreview
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ── Constantes ────────────────────────────────────────────────────────────────
$QRCoderVersion = '1.6.0'
# Hash del paquete oficial de nuget.org. Los paquetes de NuGet son inmutables,
# asi que un cambio aqui significa que la descarga fue manipulada.
$QRCoderNupkgSha256 = 'D84BFFE9DECF1FA2B8755610407959F998209029C3B6B95A1B282CEB9983A36F'
$LibDir = Join-Path $env:LOCALAPPDATA 'schoperena-win-setup\lib'

$ColoresConocidos = @{
    'negro'       = '#000000'
    'blanco'      = '#FFFFFF'
    'rojo'        = '#D32F2F'
    'verde'       = '#2E7D32'
    'azul'        = '#1565C0'
    'morado'      = '#7C3AED'
    'naranja'     = '#EF6C00'
    'cian'        = '#0097A7'
    'amarillo'    = '#F9A825'
    'rosa'        = '#C2185B'
    'gris'        = '#4B5563'
    'transparente' = 'transparent'
}

# ── Helpers de consola ────────────────────────────────────────────────────────
function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║   Generador de Codigos QR            ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Read-Requerido([string]$Prompt) {
    do {
        $v = (Read-Host "  $Prompt").Trim()
        if (-not $v) { Write-Host "  Este campo es obligatorio." -ForegroundColor Yellow }
    } while (-not $v)
    return $v
}

function Read-Opcional([string]$Prompt) {
    return (Read-Host "  $Prompt (Enter = omitir)").Trim()
}

# ── Carga de QRCoder (descarga unica, cacheada) ───────────────────────────────
function Import-QRCoder {
    if ('QRCoder.QRCodeGenerator' -as [type]) { return }

    $tfm = if ($PSVersionTable.PSEdition -eq 'Core') { 'netstandard2.0' } else { 'net40' }
    $dll = Join-Path $LibDir "QRCoder-$QRCoderVersion-$tfm.dll"

    if (-not (Test-Path $dll)) {
        Write-Host "  Descargando QRCoder $QRCoderVersion (solo la primera vez)..." -ForegroundColor Yellow

        $tmp = Join-Path ([IO.Path]::GetTempPath()) "qrcoder-$([guid]::NewGuid())"
        New-Item -ItemType Directory -Path $tmp -Force | Out-Null
        try {
            $zip = Join-Path $tmp 'qrcoder.zip'
            $url = "https://www.nuget.org/api/v2/package/QRCoder/$QRCoderVersion"
            try {
                Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing
            }
            catch {
                throw "No se pudo descargar QRCoder desde nuget.org. Revisa tu conexion a internet.`n  $($_.Exception.Message)"
            }

            $hash = (Get-FileHash $zip -Algorithm SHA256).Hash
            if ($hash -ne $QRCoderNupkgSha256) {
                throw "El paquete descargado no coincide con el hash esperado.`n  Esperado: $QRCoderNupkgSha256`n  Obtenido: $hash"
            }

            Expand-Archive -Path $zip -DestinationPath "$tmp\pkg" -Force
            $origen = Join-Path $tmp "pkg\lib\$tfm\QRCoder.dll"
            if (-not (Test-Path $origen)) { throw "El paquete no contiene lib\$tfm\QRCoder.dll" }

            New-Item -ItemType Directory -Path $LibDir -Force | Out-Null
            Copy-Item $origen $dll -Force
            Write-Host "  QRCoder guardado en $LibDir" -ForegroundColor DarkGray
        }
        finally {
            Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Add-Type -Path $dll
}

# ── Colores ───────────────────────────────────────────────────────────────────
# Devuelve un byte[4] RGBA. "transparente" -> alpha 0.
function ConvertTo-Rgba([string]$Color) {
    $c = $Color.Trim().ToLower()
    if ($ColoresConocidos.ContainsKey($c)) { $c = $ColoresConocidos[$c] }

    if ($c -in 'transparent', '#transparent') { return [byte[]](255, 255, 255, 0) }

    $hex = $c.TrimStart('#')
    if ($hex -match '^[0-9a-f]{3}$') {
        $hex = "$($hex[0])$($hex[0])$($hex[1])$($hex[1])$($hex[2])$($hex[2])"
    }
    if ($hex -notmatch '^[0-9a-f]{6}$') {
        throw "Color invalido: '$Color'. Usa #RRGGBB, #RGB o un nombre: $($ColoresConocidos.Keys -join ', ')"
    }

    return [byte[]](
        [Convert]::ToByte($hex.Substring(0, 2), 16),
        [Convert]::ToByte($hex.Substring(2, 2), 16),
        [Convert]::ToByte($hex.Substring(4, 2), 16),
        255
    )
}

function ConvertTo-Hex([byte[]]$Rgba) {
    if ($Rgba[3] -eq 0) { return 'transparent' }
    return '#{0:X2}{1:X2}{2:X2}' -f $Rgba[0], $Rgba[1], $Rgba[2]
}

# Luminancia relativa segun WCAG 2.x.
function Get-Luminancia([byte[]]$Rgba) {
    $canales = $Rgba[0..2] | ForEach-Object {
        $s = $_ / 255
        if ($s -le 0.03928) { $s / 12.92 } else { [Math]::Pow(($s + 0.055) / 1.055, 2.4) }
    }
    return 0.2126 * $canales[0] + 0.7152 * $canales[1] + 0.0722 * $canales[2]
}

# Los lectores de QR necesitan contraste alto. Avisamos si la combinacion es riesgosa.
function Test-Contraste([byte[]]$Dark, [byte[]]$Light) {
    # Un fondo transparente se vera sobre el color de la app que lo muestre; asumimos blanco.
    $lum1 = Get-Luminancia $Dark
    $lum2 = Get-Luminancia $Light
    $ratio = (([Math]::Max($lum1, $lum2)) + 0.05) / (([Math]::Min($lum1, $lum2)) + 0.05)

    Write-Host ""
    if ($ratio -lt 3) {
        Write-Host ("  AVISO: contraste bajo ({0:N1}:1). Muchos lectores no podran escanearlo." -f $ratio) -ForegroundColor Red
        Write-Host "  Se recomienda un ratio de al menos 7:1." -ForegroundColor DarkGray
    }
    elseif ($ratio -lt 7) {
        Write-Host ("  Contraste justo ({0:N1}:1). Deberia funcionar, pero pruebalo antes de imprimirlo." -f $ratio) -ForegroundColor Yellow
    }
    else {
        Write-Host ("  Contraste: {0:N1}:1 (bueno)" -f $ratio) -ForegroundColor DarkGray
    }

    if ($lum1 -gt $lum2) {
        Write-Host "  AVISO: los modulos son mas claros que el fondo (QR invertido)." -ForegroundColor Yellow
        Write-Host "  Algunos lectores antiguos fallan con QR invertidos." -ForegroundColor DarkGray
    }
}

# ── Vista previa en la terminal ───────────────────────────────────────────────
# Dibuja dos filas de modulos por linea usando el medio-bloque superior "▀":
# el color de texto pinta la fila de arriba y el de fondo la de abajo.
function Show-Preview($QrData, [byte[]]$Dark, [byte[]]$Light) {
    $m = $QrData.ModuleMatrix
    $n = $m.Count

    # Algunos hosts (ISE, procesos sin consola) no exponen el tamano de ventana.
    try { $ancho = $Host.UI.RawUI.WindowSize.Width } catch { $ancho = 0 }
    if ($ancho -gt 0 -and $n + 2 -gt $ancho) {
        Write-Host "  (vista previa omitida: la ventana es muy angosta)" -ForegroundColor DarkGray
        return
    }

    # Un fondo transparente se dibuja blanco en la terminal.
    $lz = if ($Light[3] -eq 0) { [byte[]](255, 255, 255, 255) } else { $Light }
    $e = [char]27
    $bloque = [char]0x2580
    $fgDark = "$e[38;2;$($Dark[0]);$($Dark[1]);$($Dark[2])m"
    $fgLight = "$e[38;2;$($lz[0]);$($lz[1]);$($lz[2])m"
    $bgDark = "$e[48;2;$($Dark[0]);$($Dark[1]);$($Dark[2])m"
    $bgLight = "$e[48;2;$($lz[0]);$($lz[1]);$($lz[2])m"

    Write-Host ""
    for ($y = 0; $y -lt $n; $y += 2) {
        $sb = [Text.StringBuilder]::new('  ')
        for ($x = 0; $x -lt $n; $x++) {
            $arriba = $m[$y][$x]
            $abajo = ($y + 1 -lt $n) -and $m[$y + 1][$x]
            [void]$sb.Append($(if ($arriba) { $fgDark } else { $fgLight }))
            [void]$sb.Append($(if ($abajo) { $bgDark } else { $bgLight }))
            [void]$sb.Append($bloque)
        }
        Write-Host "$($sb.ToString())$e[0m"
    }
}

# ── Constructores de payload ──────────────────────────────────────────────────
function New-PayloadUrl {
    $url = Read-Requerido "URL (ej: github.com/schoperena)"
    if ($url -notmatch '^[a-z][a-z0-9+.-]*://') { $url = "https://$url" }
    return @{ Contenido = $url; Slug = ([uri]$url).Host }
}

function New-PayloadTexto {
    $t = Read-Requerido "Texto"
    return @{ Contenido = $t; Slug = 'texto' }
}

function New-PayloadWifi {
    $ssid = Read-Requerido "Nombre de la red (SSID)"

    Write-Host ""
    Write-Host "  Seguridad:" -ForegroundColor Cyan
    Write-Host "    [1] WPA / WPA2  (lo normal)" -ForegroundColor Green
    Write-Host "    [2] WEP         (redes viejas)"
    Write-Host "    [3] Abierta     (sin contrasena)"
    Write-Host ""
    $sel = (Read-Host "  Opcion (Enter = 1)").Trim()

    $auth = switch ($sel) {
        '2' { 'WEP' }
        '3' { 'nopass' }
        default { 'WPA' }
    }

    $pass = ''
    if ($auth -ne 'nopass') {
        $sec = Read-Host "  Contrasena" -AsSecureString
        $pass = [Net.NetworkCredential]::new('', $sec).Password
        if (-not $pass) { throw "La contrasena es obligatoria para redes $auth." }
    }

    $oculta = (Read-Host "  La red esta oculta? (s/N)").Trim().ToLower() -eq 's'

    $authEnum = [QRCoder.PayloadGenerator+WiFi+Authentication]::$auth
    $p = [QRCoder.PayloadGenerator+WiFi]::new($ssid, $pass, $authEnum, $oculta, $true)
    return @{ Contenido = $p.ToString(); Slug = "wifi-$ssid"; Sensible = $true }
}

function New-PayloadContacto {
    $nombre = Read-Requerido "Nombre"
    $apellido = Read-Opcional "Apellido"
    $movil = Read-Opcional "Celular"
    $email = Read-Opcional "Email"
    $empresa = Read-Opcional "Empresa"
    $cargo = Read-Opcional "Cargo"
    $web = Read-Opcional "Sitio web"

    $tipo = [QRCoder.PayloadGenerator+ContactData+ContactOutputType]::VCard3
    $orden = [QRCoder.PayloadGenerator+ContactData+AddressOrder]::Default
    $p = [QRCoder.PayloadGenerator+ContactData]::new(
        $tipo, $nombre, $apellido,
        '',      # nickname
        '',      # phone
        $movil,
        '',      # workPhone
        $email,
        $null,   # birthday
        $web,
        '', '', '', '', '',  # street, houseNumber, city, zipCode, country
        '',      # note
        '',      # stateRegion
        $orden, $empresa, $cargo)

    return @{ Contenido = $p.ToString(); Slug = "contacto-$nombre" }
}

function New-PayloadEmail {
    $dest = Read-Requerido "Destinatario"
    $asunto = Read-Opcional "Asunto"
    $cuerpo = Read-Opcional "Mensaje"
    $enc = [QRCoder.PayloadGenerator+Mail+MailEncoding]::MAILTO
    $p = [QRCoder.PayloadGenerator+Mail]::new($dest, $asunto, $cuerpo, $enc)
    return @{ Contenido = $p.ToString(); Slug = "email-$dest" }
}

function New-PayloadSms {
    $num = Read-Requerido "Numero (ej: +573001234567)"
    $msg = Read-Opcional "Mensaje"
    $enc = [QRCoder.PayloadGenerator+SMS+SMSEncoding]::SMS
    $p = [QRCoder.PayloadGenerator+SMS]::new($num, $msg, $enc)
    return @{ Contenido = $p.ToString(); Slug = 'sms' }
}

function New-PayloadTelefono {
    $num = Read-Requerido "Numero (ej: +573001234567)"
    $p = [QRCoder.PayloadGenerator+PhoneNumber]::new($num)
    return @{ Contenido = $p.ToString(); Slug = 'tel' }
}

function New-PayloadWhatsApp {
    $num = Read-Requerido "Numero con codigo de pais (ej: 573001234567)"
    $msg = Read-Opcional "Mensaje predefinido"
    $p = [QRCoder.PayloadGenerator+WhatsAppMessage]::new($num, $msg)
    return @{ Contenido = $p.ToString(); Slug = 'whatsapp' }
}

function New-PayloadUbicacion {
    $lat = Read-Requerido "Latitud (ej: 4.710989)"
    $lon = Read-Requerido "Longitud (ej: -74.072092)"
    $enc = [QRCoder.PayloadGenerator+Geolocation+GeolocationEncoding]::GEO
    $p = [QRCoder.PayloadGenerator+Geolocation]::new($lat, $lon, $enc)
    return @{ Contenido = $p.ToString(); Slug = 'ubicacion' }
}

# ── Menus ─────────────────────────────────────────────────────────────────────
function Select-TipoContenido {
    Write-Host "  Que quieres codificar?" -ForegroundColor Cyan
    Write-Host "    [1] Enlace / URL      (por defecto)" -ForegroundColor Green
    Write-Host "    [2] Texto plano"
    Write-Host "    [3] Red WiFi"
    Write-Host "    [4] Contacto (vCard)"
    Write-Host "    [5] Email"
    Write-Host "    [6] SMS"
    Write-Host "    [7] Telefono"
    Write-Host "    [8] WhatsApp"
    Write-Host "    [9] Ubicacion (mapa)"
    Write-Host ""
    $sel = (Read-Host "  Opcion (Enter = 1)").Trim()
    Write-Host ""

    switch ($sel) {
        '2' { return New-PayloadTexto }
        '3' { return New-PayloadWifi }
        '4' { return New-PayloadContacto }
        '5' { return New-PayloadEmail }
        '6' { return New-PayloadSms }
        '7' { return New-PayloadTelefono }
        '8' { return New-PayloadWhatsApp }
        '9' { return New-PayloadUbicacion }
        default { return New-PayloadUrl }
    }
}

function Select-Colores {
    Write-Host ""
    Write-Host "  Colores:" -ForegroundColor Cyan
    Write-Host "    [1] Clasico     negro sobre blanco  (por defecto)" -ForegroundColor Green
    Write-Host "    [2] Invertido   blanco sobre negro"
    Write-Host "    [3] GitHub      azul sobre azul oscuro"
    Write-Host "    [4] Terminal    verde sobre negro"
    Write-Host "    [5] Morado      morado sobre blanco"
    Write-Host "    [6] Fondo transparente (negro, PNG)"
    Write-Host "    [7] Personalizado"
    Write-Host ""
    $sel = (Read-Host "  Opcion (Enter = 1)").Trim()

    switch ($sel) {
        '2' { return @{ Dark = '#FFFFFF'; Light = '#000000' } }
        '3' { return @{ Dark = '#58A6FF'; Light = '#0D1117' } }
        '4' { return @{ Dark = '#00E676'; Light = '#0B0F0B' } }
        '5' { return @{ Dark = '#7C3AED'; Light = '#FFFFFF' } }
        '6' { return @{ Dark = '#000000'; Light = 'transparente' } }
        '7' {
            Write-Host ""
            Write-Host "  Acepta #RRGGBB, #RGB o nombres: $($ColoresConocidos.Keys -join ', ')" -ForegroundColor DarkGray
            $d = (Read-Host "  Color de los modulos (Enter = negro)").Trim()
            $l = (Read-Host "  Color del fondo    (Enter = blanco)").Trim()
            return @{
                Dark  = $(if ($d) { $d } else { 'negro' })
                Light = $(if ($l) { $l } else { 'blanco' })
            }
        }
        default { return @{ Dark = '#000000'; Light = '#FFFFFF' } }
    }
}

function Select-Formato {
    Write-Host ""
    Write-Host "  Formato:" -ForegroundColor Cyan
    Write-Host "    [1] PNG  imagen, para compartir o imprimir  (por defecto)" -ForegroundColor Green
    Write-Host "    [2] SVG  vectorial, escala sin perder calidad"
    Write-Host ""
    $sel = (Read-Host "  Opcion (Enter = 1)").Trim()
    return $(if ($sel -eq '2') { 'svg' } else { 'png' })
}

# ── Generacion ────────────────────────────────────────────────────────────────
function ConvertTo-Slug([string]$Texto) {
    $s = $Texto -replace '[^\w\-]+', '-' -replace '-{2,}', '-'
    $s = $s.Trim('-').ToLower()
    if ($s.Length -gt 40) { $s = $s.Substring(0, 40).Trim('-') }
    return $(if ($s) { $s } else { 'qr' })
}

function New-QRFile {
    param(
        [string]   $Contenido,
        [string]   $Ruta,
        [byte[]]   $DarkRgba,
        [byte[]]   $LightRgba,
        [int]      $Escala,
        [string]   $Ecc,
        [string]   $Formato,
        [bool]     $DibujarMargen
    )

    $gen = [QRCoder.QRCodeGenerator]::new()
    $eccEnum = [QRCoder.QRCodeGenerator+ECCLevel]::$Ecc
    $data = $gen.CreateQrCode($Contenido, $eccEnum)

    if ($Formato -eq 'svg') {
        $svg = [QRCoder.SvgQRCode]::new($data)
        $modo = [QRCoder.SvgQRCode+SizingMode]::WidthHeightAttribute
        $txt = $svg.GetGraphic($Escala, (ConvertTo-Hex $DarkRgba), (ConvertTo-Hex $LightRgba), $DibujarMargen, $modo, $null)
        [IO.File]::WriteAllText($Ruta, $txt, [Text.UTF8Encoding]::new($false))
    }
    else {
        $png = [QRCoder.PngByteQRCode]::new($data)
        $bytes = $png.GetGraphic($Escala, $DarkRgba, $LightRgba, $DibujarMargen)
        [IO.File]::WriteAllBytes($Ruta, $bytes)
    }

    return $data
}

# ── Main ──────────────────────────────────────────────────────────────────────
$interactivo = -not $Text
$sensible = $false

if ($interactivo) { Write-Header } else { Write-Host "" }

Import-QRCoder

if ($interactivo) {
    $payload = Select-TipoContenido
    $contenido = $payload.Contenido
    $slug = ConvertTo-Slug $payload.Slug
    $sensible = [bool]$payload.Sensible

    # Si ya vinieron por parametro, no volvemos a preguntar.
    if (-not ($PSBoundParameters.ContainsKey('Dark') -or $PSBoundParameters.ContainsKey('Light'))) {
        $colores = Select-Colores
        $Dark = $colores.Dark
        $Light = $colores.Light
    }
    if (-not $PSBoundParameters.ContainsKey('Formato')) {
        $Formato = Select-Formato
    }

    if ($Light -eq 'transparente' -and $Formato -eq 'svg') {
        Write-Host ""
        Write-Host "  Nota: el fondo transparente en SVG depende del visor." -ForegroundColor DarkGray
    }
}
else {
    $contenido = $Text
    $slug = ConvertTo-Slug $(if ($Text -match '^[a-z]+://([^/\s]+)') { $matches[1] } else { $Text })
}

$darkRgba = ConvertTo-Rgba $Dark
$lightRgba = ConvertTo-Rgba $Light
Test-Contraste $darkRgba $lightRgba

if (-not $Out) {
    $escritorio = [Environment]::GetFolderPath('Desktop')
    $carpeta = $(if ($escritorio -and (Test-Path $escritorio)) { $escritorio } else { $PWD.Path })
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $Out = Join-Path $carpeta "qr-$slug-$stamp.$Formato"
}
elseif (-not [IO.Path]::IsPathRooted($Out)) {
    $Out = Join-Path $PWD.Path $Out
}

$destino = [IO.Path]::GetDirectoryName($Out)
if ($destino -and -not (Test-Path $destino)) {
    New-Item -ItemType Directory -Path $destino -Force | Out-Null
}

$data = New-QRFile -Contenido $contenido -Ruta $Out `
    -DarkRgba $darkRgba -LightRgba $lightRgba `
    -Escala $Escala -Ecc $Ecc -Formato $Formato -DibujarMargen (-not $SinMargen)

if (-not $SinPreview) { Show-Preview $data $darkRgba $lightRgba }

$info = Get-Item $Out
$peso = if ($info.Length -lt 1KB) { "$($info.Length) bytes" } else { '{0:N1} KB' -f ($info.Length / 1KB) }

Write-Host ""
Write-Host "  QR generado:" -ForegroundColor Green
Write-Host "    Archivo   : $Out"
Write-Host "    Tamano    : $peso"
Write-Host "    Modulos   : $($data.ModuleMatrix.Count) x $($data.ModuleMatrix.Count) (con margen)"
Write-Host "    Correccion: $Ecc"

# No imprimimos contrasenas de WiFi en pantalla.
if (-not $sensible) {
    $vista = $(if ($contenido.Length -gt 70) { $contenido.Substring(0, 70) + '...' } else { $contenido })
    Write-Host "    Contenido : $($vista -replace '\r?\n', ' | ')" -ForegroundColor DarkGray
}

if ($interactivo) {
    Write-Host ""
    $abrir = (Read-Host "  Abrir el archivo? (S/n)").Trim().ToLower()
    if ($abrir -ne 'n') { Invoke-Item $Out }
    Write-Host ""
    Pause
}
