<#
.SYNOPSIS
    Divide un PDF grande (cartel, plano, poster) en varias hojas imprimibles.

.DESCRIPTION
    Recorta el PDF en una cuadricula de paginas del tamano de papel que elijas
    (A4 por defecto). Cada hoja sale exactamente del tamano del papel, sin
    deformar el contenido, asi que al imprimir y pegar las hojas el cartel
    queda a escala real.

    Sin argumentos abre un menu: arrastras el PDF a la ventana, te muestra el
    tamano detectado del original y eliges el papel de destino.

    Usa PyMuPDF; si no esta instalado lo instala con pip automaticamente.

.EXAMPLE
    .\DividirPDF.ps1
    Menu interactivo.

.EXAMPLE
    .\DividirPDF.ps1 '.\cartel venta.pdf'
    Divide el cartel en hojas A4 al tamano original, sin preguntar nada.

.EXAMPLE
    .\DividirPDF.ps1 .\cartel.pdf -Escala 2 -Solape 10 -Marcas
    Agranda el cartel al doble, deja 10 mm de solape para pegar y dibuja
    lineas de corte con etiquetas de posicion.

.EXAMPLE
    .\DividirPDF.ps1 .\cartel.pdf -Columnas 4 -Filas 2
    Fuerza una cuadricula de 4x2 en vez de calcularla automaticamente.
#>

[CmdletBinding()]
param(
    # PDF de entrada. Si se omite, se abre el menu interactivo.
    [Parameter(Position = 0)]
    [string] $Ruta,

    # PDF de salida. Por defecto: <entrada>-mosaico.pdf junto al original.
    [string] $Salida,

    # Tamano de cada hoja resultante.
    [ValidateSet('A3', 'A4', 'A5', 'Carta', 'Oficio', 'Tabloide')]
    [string] $Tamano = 'A4',

    # Factor de ampliacion del cartel antes de recortarlo. 2 = el doble de grande.
    [ValidateRange(0.05, 50)]
    [double] $Escala = 1.0,

    # Milimetros que se repiten entre hojas vecinas, para tener pestana al pegar.
    [ValidateRange(0, 100)]
    [double] $Solape = 0,

    # Margen no imprimible de tu impresora, en milimetros.
    [ValidateRange(0, 50)]
    [double] $Margen = 0,

    # Fuerza la cuadricula en vez de calcularla. Requiere ambos.
    [ValidateRange(0, 50)]
    [int] $Columnas = 0,

    [ValidateRange(0, 50)]
    [int] $Filas = 0,

    # Pagina del PDF de origen a dividir (1 = la primera).
    [ValidateRange(1, 10000)]
    [int] $Pagina = 1,

    # Dibuja lineas de corte punteadas y etiqueta cada hoja (fila,columna).
    [switch] $Marcas,

    # No abre el PDF resultante al terminar.
    [switch] $NoAbrir
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Puntos PostScript (1 pt = 1/72 pulgada), siempre en vertical.
$Papeles = [ordered]@{
    'A5'       = @(419.53, 595.28)
    'A4'       = @(595.28, 841.89)
    'A3'       = @(841.89, 1190.55)
    'Carta'    = @(612.0, 792.0)
    'Oficio'   = @(612.0, 1008.0)
    'Tabloide' = @(792.0, 1224.0)
}

# Solo para reconocer el tamano del PDF de origen; no se puede imprimir en ellos.
$PapelesGrandes = [ordered]@{
    'A2' = @(1190.55, 1683.78)
    'A1' = @(1683.78, 2383.94)
    'A0' = @(2383.94, 3370.39)
}

$MM = 72 / 25.4

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║   Dividir PDF en hojas imprimibles   ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Abortar([string]$Mensaje) {
    Write-Host ""
    Write-Host "  ERROR: $Mensaje" -ForegroundColor Red
    Write-Host ""
    Pause
    exit 1
}

# ── 1. Python ─────────────────────────────────────────────────────────────────
function Get-Python {
    foreach ($c in 'python', 'py', 'python3') {
        $cmd = Get-Command $c -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        # El alias de la Microsoft Store existe pero no ejecuta nada hasta instalarse.
        $v = & $c --version 2>&1
        if ($LASTEXITCODE -eq 0 -and $v -match 'Python 3') { return $c }
    }
    return $null
}

function Install-PyMuPDF([string]$Py) {
    & $Py -c "import pymupdf" 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { return }

    # Las versiones viejas solo exponen el modulo como 'fitz'.
    & $Py -c "import fitz" 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { return }

    Write-Host "  PyMuPDF no esta instalado. Instalando con pip..." -ForegroundColor Yellow
    Write-Host ""
    & $Py -m pip install --upgrade pymupdf
    if ($LASTEXITCODE -ne 0) {
        Abortar "fallo 'pip install pymupdf'. Revisa tu instalacion de Python."
    }
    Write-Host ""
}

# ── 2. Script de Python embebido ──────────────────────────────────────────────
# Tres modos: 'info' lee el PDF, 'calcular' resuelve la cuadricula sin escribir
# nada (para la vista previa del menu) y 'generar' produce el mosaico.
$PythonSource = @'
import sys, json, math

try:
    import pymupdf as fitz
except ImportError:
    import fitz

cfg = json.load(open(sys.argv[1], encoding="utf-8"))
MM = 72 / 25.4

doc = fitz.open(cfg["entrada"])

if cfg["modo"] == "info":
    print(json.dumps({
        "paginas": len(doc),
        "tamanos": [[round(p.rect.width, 2), round(p.rect.height, 2)] for p in doc],
    }))
    sys.exit(0)

if cfg["pagina"] > len(doc):
    sys.exit(f"ERROR: el PDF tiene {len(doc)} pagina(s); pediste la {cfg['pagina']}.")

margen = cfg["margen"] * MM
solape = cfg["solape"] * MM
escala = cfg["escala"]
pno    = cfg["pagina"] - 1
src    = doc[pno].rect
outW   = src.width  * escala
outH   = src.height * escala


def cuadricula(pw, ph):
    """Cuantas hojas de pw x ph se necesitan para cubrir el cartel."""
    cw, ch = pw - 2 * margen, ph - 2 * margen
    paso_x, paso_y = cw - solape, ch - solape
    if paso_x <= 1 or paso_y <= 1:
        return None
    cols = max(1, math.ceil((outW - solape) / paso_x))
    rows = max(1, math.ceil((outH - solape) / paso_y))
    return cols, rows, pw, ph, cw, ch


pw, ph = cfg["papelW"], cfg["papelH"]

if cfg["cols"] and cfg["rows"]:
    # Cuadricula forzada: la hoja se orienta segun la forma del recorte.
    cols, rows = cfg["cols"], cfg["rows"]
    if (outW / cols > outH / rows) != (pw > ph):
        pw, ph = ph, pw
    cw, ch = pw - 2 * margen, ph - 2 * margen
    if cw <= 1 or ch <= 1:
        sys.exit("ERROR: el margen es mas grande que la hoja.")
    auto = False
else:
    # Probamos la hoja en vertical y en horizontal, y nos quedamos con la que
    # gaste menos paginas (a igualdad, la que desperdicie menos papel).
    opciones = [c for c in (cuadricula(pw, ph), cuadricula(ph, pw)) if c]
    if not opciones:
        sys.exit("ERROR: el solape o el margen no dejan area util en la hoja.")
    cols, rows, pw, ph, cw, ch = min(
        opciones, key=lambda o: (o[0] * o[1], o[0] * o[4] * o[1] * o[5]))
    auto = True

resumen = {
    "cols": cols, "rows": rows, "paginas": cols * rows,
    "orientacion": "horizontal" if pw > ph else "vertical",
    "cartelW": round(outW / MM, 1), "cartelH": round(outH / MM, 1),
}

if cfg["modo"] == "calcular":
    print(json.dumps(resumen))
    sys.exit(0)

out = fitz.open()

for r in range(rows):
    for c in range(cols):
        if auto:
            # Cada hoja cubre cw x ch del cartel; las vecinas se solapan.
            x0 = c * (cw - solape)
            y0 = r * (ch - solape)
            x1 = min(x0 + cw, outW)
            y1 = min(y0 + ch, outH)
        else:
            x0, y0 = c * outW / cols, r * outH / rows
            x1, y1 = (c + 1) * outW / cols, (r + 1) * outH / rows

        # Del espacio de salida (ya escalado) al del PDF original.
        clip = fitz.Rect(x0 / escala, y0 / escala, x1 / escala, y1 / escala) & src
        if clip.is_empty:
            continue

        ancho, alto = clip.width * escala, clip.height * escala
        pagina = out.new_page(width=pw, height=ph)

        if auto:
            destino = fitz.Rect(margen, margen, margen + ancho, margen + alto)
        else:
            # Cuadricula forzada: encajar sin deformar, centrado.
            f = min(cw / ancho, ch / alto)
            aw, ah = ancho * f, alto * f
            ox, oy = (pw - aw) / 2, (ph - ah) / 2
            destino = fitz.Rect(ox, oy, ox + aw, oy + ah)

        pagina.show_pdf_page(destino, doc, pno, clip=clip)

        if cfg["marcas"]:
            gris = (0.6, 0.6, 0.6)
            pagina.draw_rect(destino, color=gris, width=0.4, dashes="[3 3] 0")
            pagina.insert_text((margen + 4, ph - 6), f"fila {r + 1} / col {c + 1}",
                               fontsize=7, color=gris)

out.save(cfg["salida"], garbage=4, deflate=True)
resumen["paginas"] = out.page_count
print(json.dumps(resumen))
'@

# ── 3. Puente a Python ────────────────────────────────────────────────────────
function Invoke-Py([string]$Py, [hashtable]$Cfg) {
    $tmp = Join-Path ([IO.Path]::GetTempPath()) "dividirpdf-$([guid]::NewGuid())"
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    try {
        $cfgFile = Join-Path $tmp 'cfg.json'
        $pyFile = Join-Path $tmp 'dividir.py'
        $utf8 = [Text.UTF8Encoding]::new($false)
        [IO.File]::WriteAllText($cfgFile, ($Cfg | ConvertTo-Json -Compress), $utf8)
        [IO.File]::WriteAllText($pyFile, $PythonSource, $utf8)

        $env:PYTHONIOENCODING = 'utf-8'
        $salida = & $Py $pyFile $cfgFile 2>&1
        if ($LASTEXITCODE -ne 0) {
            Abortar ($salida -join "`n  ")
        }
        return ($salida | Select-Object -Last 1 | ConvertFrom-Json)
    }
    finally {
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ── 4. Deteccion del tamano de origen ─────────────────────────────────────────
# Reconoce el papel del PDF de entrada, en cualquiera de las dos orientaciones.
function Get-NombrePapel([double]$W, [double]$H) {
    $todos = @{}
    foreach ($k in $Papeles.Keys) { $todos[$k] = $Papeles[$k] }
    foreach ($k in $PapelesGrandes.Keys) { $todos[$k] = $PapelesGrandes[$k] }

    foreach ($nombre in $todos.Keys) {
        $p = $todos[$nombre]
        $vertical = ([Math]::Abs($W - $p[0]) -lt 3) -and ([Math]::Abs($H - $p[1]) -lt 3)
        $apaisado = ([Math]::Abs($W - $p[1]) -lt 3) -and ([Math]::Abs($H - $p[0]) -lt 3)
        if ($vertical) { return "$nombre vertical" }
        if ($apaisado) { return "$nombre horizontal" }
    }
    return $null
}

function Show-Origen([double]$W, [double]$H, [int]$Paginas, [int]$Pag) {
    $nombre = Get-NombrePapel $W $H
    $orient = if ($W -gt $H) { 'horizontal' } elseif ($H -gt $W) { 'vertical' } else { 'cuadrado' }

    Write-Host ""
    Write-Host "  Origen detectado:" -ForegroundColor Cyan
    Write-Host ("    Tamano  : {0:N1} x {1:N1} mm  ({2:N1} x {3:N1} cm)" -f `
        ($W / $MM), ($H / $MM), ($W / $MM / 10), ($H / $MM / 10))
    if ($nombre) {
        Write-Host "    Formato : $nombre" -ForegroundColor Green
    }
    else {
        Write-Host "    Formato : no estandar ($orient)" -ForegroundColor DarkGray
    }
    if ($Paginas -gt 1) {
        Write-Host "    Paginas : $Paginas (se dividira la $Pag)"
    }
}

# ── 5. Menu ───────────────────────────────────────────────────────────────────
# Al arrastrar un archivo a la terminal, Windows envuelve la ruta en comillas
# si tiene espacios. Tambien limpiamos comillas simples y espacios sobrantes.
function Read-RutaPdf {
    Write-Host "  Arrastra el PDF a esta ventana y pulsa Enter," -ForegroundColor Cyan
    Write-Host "  o pega la ruta completa. (q = salir)" -ForegroundColor DarkGray
    Write-Host ""

    while ($true) {
        $entrada = (Read-Host "  PDF").Trim()
        if ($entrada -in 'q', 'Q') { Write-Host ""; exit 0 }
        if (-not $entrada) { continue }

        $limpia = $entrada.Trim('"', "'", ' ')
        $r = Resolve-Path -LiteralPath $limpia -ErrorAction SilentlyContinue
        if (-not $r) {
            Write-Host "  No se encontro: $limpia" -ForegroundColor Yellow
            continue
        }
        if ([IO.Path]::GetExtension($r.Path) -ne '.pdf') {
            Write-Host "  No es un PDF: $([IO.Path]::GetFileName($r.Path))" -ForegroundColor Yellow
            continue
        }
        return $r.Path
    }
}

function Select-Pagina([int]$Total) {
    if ($Total -le 1) { return 1 }
    Write-Host ""
    while ($true) {
        $s = (Read-Host "  Que pagina quieres dividir? (Enter = 1)").Trim()
        if (-not $s) { return 1 }
        if ($s -match '^\d+$' -and [int]$s -ge 1 -and [int]$s -le $Total) { return [int]$s }
        Write-Host "  Escribe un numero entre 1 y $Total." -ForegroundColor Yellow
    }
}

function Select-Destino {
    Write-Host ""
    Write-Host "  Papel de destino (cada hoja saldra de este tamano):" -ForegroundColor Cyan
    $nombres = @($Papeles.Keys)
    for ($i = 0; $i -lt $nombres.Count; $i++) {
        $n = $nombres[$i]
        $p = $Papeles[$n]
        $etq = '{0,-9} {1,3:N0} x {2,3:N0} mm' -f $n, ($p[0] / $MM), ($p[1] / $MM)
        if ($n -eq 'A4') {
            Write-Host "    [$($i + 1)] $etq  (por defecto)" -ForegroundColor Green
        }
        else {
            Write-Host "    [$($i + 1)] $etq"
        }
    }
    Write-Host ""
    $s = (Read-Host "  Opcion (Enter = A4)").Trim()
    if (-not $s) { return 'A4' }
    if ($s -match '^\d+$' -and [int]$s -ge 1 -and [int]$s -le $nombres.Count) {
        return $nombres[[int]$s - 1]
    }
    # Tambien aceptamos el nombre escrito directo.
    foreach ($n in $nombres) { if ($s -eq $n) { return $n } }
    Write-Host "  Opcion invalida, uso A4." -ForegroundColor Yellow
    return 'A4'
}

function Read-Numero([string]$Prompt, [double]$Defecto, [double]$Min, [double]$Max) {
    while ($true) {
        $s = (Read-Host "  $Prompt").Trim().Replace(',', '.')
        if (-not $s) { return $Defecto }
        $v = 0.0
        $ok = [double]::TryParse($s, [Globalization.NumberStyles]::Float,
            [Globalization.CultureInfo]::InvariantCulture, [ref]$v)
        if ($ok -and $v -ge $Min -and $v -le $Max) { return $v }
        Write-Host "  Escribe un numero entre $Min y $Max." -ForegroundColor Yellow
    }
}

# ── 6. Main ───────────────────────────────────────────────────────────────────
if (($Columnas -gt 0) -xor ($Filas -gt 0)) {
    Abortar "usa -Columnas y -Filas juntos, o ninguno de los dos."
}

$interactivo = -not $Ruta
Write-Header

$py = Get-Python
if (-not $py) {
    Write-Host ""
    Write-Host "  ERROR: no se encontro Python 3." -ForegroundColor Red
    Write-Host "  Instalalo con: winget install Python.Python.3.12" -ForegroundColor Yellow
    Write-Host ""
    Pause; exit 1
}
Install-PyMuPDF $py

# --- 6.1 Ruta del PDF ---
if ($interactivo) {
    $Ruta = Read-RutaPdf
}
else {
    $r = Resolve-Path -LiteralPath $Ruta -ErrorAction SilentlyContinue
    if (-not $r) { Abortar "no se encontro el archivo: $Ruta" }
    $Ruta = $r.Path
    if ([IO.Path]::GetExtension($Ruta) -ne '.pdf') {
        Abortar "el archivo debe ser un PDF: $([IO.Path]::GetFileName($Ruta))"
    }
}

# --- 6.2 Origen: leer y mostrar lo detectado ---
$info = Invoke-Py $py @{ modo = 'info'; entrada = $Ruta }
if ($Pagina -gt $info.paginas) {
    Abortar "el PDF tiene $($info.paginas) pagina(s); pediste la $Pagina."
}

if ($interactivo) {
    $t = $info.tamanos[0]
    Show-Origen $t[0] $t[1] $info.paginas 1
    $Pagina = Select-Pagina $info.paginas
    if ($Pagina -ne 1) {
        $t = $info.tamanos[$Pagina - 1]
        Show-Origen $t[0] $t[1] $info.paginas $Pagina
    }
}

# --- 6.3 Destino ---
if ($interactivo) {
    if (-not $PSBoundParameters.ContainsKey('Tamano')) { $Tamano = Select-Destino }

    Write-Host ""
    if (-not $PSBoundParameters.ContainsKey('Escala')) {
        Write-Host "  Escala: 1 = tamano original, 2 = el doble de grande." -ForegroundColor DarkGray
        $Escala = Read-Numero "Escala (Enter = 1)" 1.0 0.05 50
    }
    if (-not $PSBoundParameters.ContainsKey('Solape')) {
        Write-Host "  Solape: milimetros repetidos entre hojas, para tener pestana al pegar." -ForegroundColor DarkGray
        $Solape = Read-Numero "Solape en mm (Enter = 0)" 0 0 100
    }
    if (-not $PSBoundParameters.ContainsKey('Marcas')) {
        $Marcas = (Read-Host "  Dibujar lineas de corte y etiquetas? (s/N)").Trim().ToLower() -eq 's'
    }
}

$papel = $Papeles[$Tamano]
$cfgBase = @{
    entrada = $Ruta
    papelW  = $papel[0]
    papelH  = $papel[1]
    escala  = $Escala
    solape  = $Solape
    margen  = $Margen
    cols    = $Columnas
    rows    = $Filas
    pagina  = $Pagina
    marcas  = [bool]$Marcas
}

# --- 6.4 Vista previa y confirmacion ---
if ($interactivo) {
    $prev = Invoke-Py $py ($cfgBase + @{ modo = 'calcular'; salida = '' })
    Write-Host ""
    Write-Host "  Resultado:" -ForegroundColor Cyan
    Write-Host "    $($prev.cols) x $($prev.rows) = $($prev.paginas) hojas $Tamano $($prev.orientacion)" -ForegroundColor Green
    Write-Host ("    Cartel final: {0} x {1} mm" -f $prev.cartelW, $prev.cartelH)
    Write-Host ""
    if ((Read-Host "  Continuar? (S/n)").Trim().ToLower() -eq 'n') {
        Write-Host "  Cancelado." -ForegroundColor DarkGray
        Write-Host ""
        exit 0
    }
}

# --- 6.5 Archivo de salida ---
if (-not $Salida) {
    $dir = [IO.Path]::GetDirectoryName($Ruta)
    $nom = [IO.Path]::GetFileNameWithoutExtension($Ruta)
    $porDefecto = Join-Path $dir "$nom-mosaico.pdf"

    if ($interactivo) {
        Write-Host ""
        Write-Host "  Se guardara en: $porDefecto" -ForegroundColor DarkGray
        $s = (Read-Host "  Otra ruta? (Enter = usar esa)").Trim().Trim('"', "'", ' ')
        $Salida = if ($s) { $s } else { $porDefecto }
    }
    else {
        $Salida = $porDefecto
    }
}
if (-not [IO.Path]::IsPathRooted($Salida)) { $Salida = Join-Path $PWD.Path $Salida }
if ([IO.Path]::GetExtension($Salida) -ne '.pdf') { $Salida += '.pdf' }

if ((Resolve-Path -LiteralPath $Salida -ErrorAction SilentlyContinue).Path -eq $Ruta) {
    Abortar "el PDF de salida no puede ser el mismo que el de entrada."
}
$destino = [IO.Path]::GetDirectoryName($Salida)
if ($destino -and -not (Test-Path $destino)) {
    New-Item -ItemType Directory -Path $destino -Force | Out-Null
}

# --- 6.6 Generar ---
Write-Host ""
Write-Host "  Procesando $([IO.Path]::GetFileName($Ruta))..." -ForegroundColor Yellow
$r = Invoke-Py $py ($cfgBase + @{ modo = 'generar'; salida = $Salida })

$peso = '{0:N1} KB' -f ((Get-Item $Salida).Length / 1KB)

Write-Host ""
Write-Host "  Listo!" -ForegroundColor Green
Write-Host "    Archivo   : $Salida"
Write-Host "    Cuadricula: $($r.cols) x $($r.rows) = $($r.paginas) hojas $Tamano $($r.orientacion)"
Write-Host ("    Cartel    : {0} x {1} mm" -f $r.cartelW, $r.cartelH)
Write-Host "    Peso      : $peso"
if ($Solape -gt 0) { Write-Host "    Solape    : $Solape mm entre hojas" }
Write-Host ""
Write-Host "  Al imprimir elige 'Tamano real' o 100%, NO 'Ajustar a la pagina'," -ForegroundColor DarkGray
Write-Host "  o las hojas no encajaran entre si." -ForegroundColor DarkGray

if (-not $NoAbrir) { Invoke-Item $Salida }

Write-Host ""
Pause
