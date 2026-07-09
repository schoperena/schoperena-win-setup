<#
.SYNOPSIS
    Divide un PDF grande (cartel, plano, poster) en varias hojas imprimibles.

.DESCRIPTION
    Recorta el PDF en una cuadricula de paginas del tamano de papel que elijas
    (A4 por defecto). Cada hoja sale exactamente del tamano del papel, sin
    deformar el contenido, asi que al imprimir y pegar las hojas el cartel
    queda a escala real.

    Usa PyMuPDF; si no esta instalado lo instala con pip automaticamente.

.EXAMPLE
    .\DividirPDF.ps1 '.\cartel venta.pdf'
    Divide el cartel en hojas A4 al tamano original.

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
    # PDF de entrada.
    [Parameter(Mandatory, Position = 0)]
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

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║   Dividir PDF en hojas imprimibles   ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
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
        throw "Fallo 'pip install pymupdf'. Revisa tu instalacion de Python."
    }
    Write-Host ""
}

# ── 2. Script de Python embebido ──────────────────────────────────────────────
$PythonSource = @'
import sys, json, math

try:
    import pymupdf as fitz
except ImportError:
    import fitz

cfg = json.load(open(sys.argv[1], encoding="utf-8"))

MM = 72 / 25.4
margen  = cfg["margen"]  * MM
solape  = cfg["solape"]  * MM
escala  = cfg["escala"]

doc = fitz.open(cfg["entrada"])
if cfg["pagina"] > len(doc):
    sys.exit(f"ERROR: el PDF tiene {len(doc)} pagina(s); pediste la {cfg['pagina']}.")

pno  = cfg["pagina"] - 1
src  = doc[pno].rect
outW = src.width  * escala
outH = src.height * escala

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
        clip = fitz.Rect(x0 / escala, y0 / escala, x1 / escala, y1 / escala)
        clip = clip & src
        if clip.is_empty:
            continue

        ancho  = clip.width  * escala
        alto   = clip.height * escala
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
            etiqueta = f"fila {r + 1} / col {c + 1}"
            pagina.insert_text((margen + 4, ph - 6), etiqueta,
                               fontsize=7, color=gris)

out.save(cfg["salida"], garbage=4, deflate=True)

print(json.dumps({
    "cols": cols, "rows": rows, "paginas": out.page_count,
    "orientacion": "horizontal" if pw > ph else "vertical",
    "cartelW": round(outW / MM, 1), "cartelH": round(outH / MM, 1),
}))
'@

# ── 3. Main ───────────────────────────────────────────────────────────────────
function Abortar([string]$Mensaje) {
    Write-Host ""
    Write-Host "  ERROR: $Mensaje" -ForegroundColor Red
    Write-Host ""
    Pause
    exit 1
}

Write-Header

if (($Columnas -gt 0) -xor ($Filas -gt 0)) {
    Abortar "usa -Columnas y -Filas juntos, o ninguno de los dos."
}

$resuelta = Resolve-Path -LiteralPath $Ruta -ErrorAction SilentlyContinue
if (-not $resuelta) { Abortar "no se encontro el archivo: $Ruta" }
$Ruta = $resuelta.Path

if ([IO.Path]::GetExtension($Ruta) -ne '.pdf') {
    Abortar "el archivo debe ser un PDF: $([IO.Path]::GetFileName($Ruta))"
}

if (-not $Salida) {
    $dir = [IO.Path]::GetDirectoryName($Ruta)
    $nom = [IO.Path]::GetFileNameWithoutExtension($Ruta)
    $Salida = Join-Path $dir "$nom-mosaico.pdf"
}
elseif (-not [IO.Path]::IsPathRooted($Salida)) {
    $Salida = Join-Path $PWD.Path $Salida
}

if ((Resolve-Path -LiteralPath $Salida -ErrorAction SilentlyContinue).Path -eq $Ruta) {
    Abortar "el PDF de salida no puede ser el mismo que el de entrada."
}

$py = Get-Python
if (-not $py) {
    Write-Host ""
    Write-Host "  ERROR: no se encontro Python 3." -ForegroundColor Red
    Write-Host "  Instalalo con: winget install Python.Python.3.12" -ForegroundColor Yellow
    Write-Host ""
    Pause; exit 1
}
Install-PyMuPDF $py

# Puntos PostScript (1 pt = 1/72 pulgada), siempre en vertical.
$papeles = @{
    'A3'       = @(841.89, 1190.55)
    'A4'       = @(595.28, 841.89)
    'A5'       = @(419.53, 595.28)
    'Carta'    = @(612.0, 792.0)
    'Oficio'   = @(612.0, 1008.0)
    'Tabloide' = @(792.0, 1224.0)
}
$papel = $papeles[$Tamano]

$cfg = @{
    entrada = $Ruta
    salida  = $Salida
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

$tmp = Join-Path ([IO.Path]::GetTempPath()) "dividirpdf-$([guid]::NewGuid())"
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    $cfgFile = Join-Path $tmp 'cfg.json'
    $pyFile = Join-Path $tmp 'dividir.py'
    $utf8 = [Text.UTF8Encoding]::new($false)
    [IO.File]::WriteAllText($cfgFile, ($cfg | ConvertTo-Json -Compress), $utf8)
    [IO.File]::WriteAllText($pyFile, $PythonSource, $utf8)

    Write-Host "  Procesando $([IO.Path]::GetFileName($Ruta))..." -ForegroundColor Yellow

    $env:PYTHONIOENCODING = 'utf-8'
    $salidaPy = & $py $pyFile $cfgFile 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "  $($salidaPy -join "`n  ")" -ForegroundColor Red
        Write-Host ""
        Pause; exit 1
    }
}
finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

$r = $salidaPy | Select-Object -Last 1 | ConvertFrom-Json
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
