#Requires -Version 7
param([string]$RepoBase = $PSScriptRoot)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
try { [Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch {}

$ErrorActionPreference = 'Stop'

# ── Paleta dark ───────────────────────────────────────────────────────────────
$cBg      = [Drawing.Color]::FromArgb(18,  18,  30)
$cPanel   = [Drawing.Color]::FromArgb(26,  26,  42)
$cCard    = [Drawing.Color]::FromArgb(34,  34,  54)
$cBorder  = [Drawing.Color]::FromArgb(55,  55,  88)
$cAccent  = [Drawing.Color]::FromArgb(0,  180, 216)
$cTxt     = [Drawing.Color]::FromArgb(218, 218, 232)
$cTxtDim  = [Drawing.Color]::FromArgb(130, 130, 155)
$cGreen   = [Drawing.Color]::FromArgb(72,  199, 116)
$cYellow  = [Drawing.Color]::FromArgb(255, 200,  70)
$cRed     = [Drawing.Color]::FromArgb(255,  88,  88)
$cRowAlt  = [Drawing.Color]::FromArgb(40,  40,  62)
$cSelBg   = [Drawing.Color]::FromArgb(0,  100, 160)
$cBtnAll  = [Drawing.Color]::FromArgb(50,  50,  78)
$cBtnRun  = [Drawing.Color]::FromArgb(0,  148,  80)
$cBtnCan  = [Drawing.Color]::FromArgb(120,  30,  30)

# ── Fuentes ───────────────────────────────────────────────────────────────────
$fTitle  = New-Object Drawing.Font('Segoe UI', 16, [Drawing.FontStyle]::Bold)
$fHead   = New-Object Drawing.Font('Segoe UI', 10, [Drawing.FontStyle]::Bold)
$fNorm   = New-Object Drawing.Font('Segoe UI',  9)
$fSmall  = New-Object Drawing.Font('Segoe UI',  8)
$fMono   = New-Object Drawing.Font('Consolas',  9)

# ── Datos ─────────────────────────────────────────────────────────────────────
$browserItems = @(
    @{ label='Google Chrome';   id='Google.Chrome';               type='winget' }
    @{ label='Brave';           id='Brave.Brave';                  type='winget' }
    @{ label='Mozilla Firefox'; id='Mozilla.Firefox';              type='winget' }
    @{ label='LibreWolf';       id='LibreWolf.LibreWolf';          type='winget' }
)
$aiItems = @(
    @{ label='Claude Desktop';  id='Anthropic.Claude';            type='winget' }
    @{ label='Claude Code CLI'; id='@anthropic-ai/claude-code';   type='npm'    }
    @{ label='Codex CLI';       id='@openai/codex';               type='npm'    }
)
$scriptItems = @(
    @{ name='MenuScripts.ps1';                desc='HUB central del toolbox (recomendado)' }
    @{ name='RenombrarMasivo.ps1';            desc='Renombrado masivo con criterios múltiples y opción de revertir' }
    @{ name='BloquearAdobe.ps1';              desc='Bloquea Adobe vía archivo hosts' }
    @{ name='FormatearDisco.ps1';             desc='Formatea discos interactivamente — requiere Admin' }
    @{ name='New-SSHKey.ps1';                 desc='Genera llaves SSH Ed25519 / RSA 4096' }
    @{ name='New-QRCode.ps1';                 desc='Genera códigos QR con colores personalizables' }
    @{ name='verify-checksum.ps1';            desc='Verifica integridad via MD5 / SHA256 / SHA512' }
    @{ name='tree.ps1';                       desc='Árbol de directorios en consola' }
    @{ name='win11_rpd_patch.ps1';            desc='Habilita Escritorio Remoto (RDP) en Windows 11 Home' }
    @{ name='stirling-sch.ps1';               desc='Levanta Stirling PDF localmente vía Docker' }
    @{ name='deblotear_TCL10L.ps1';           desc='Elimina bloatware del TCL 10L vía ADB' }
    @{ name='calc_digito_de_verificacion.py'; desc='Calcula dígito de verificación NIT (Colombia)' }
    @{ name='procesar_notebook.py';           desc='Convierte y procesa Jupyter Notebooks' }
)

# ── Layout ────────────────────────────────────────────────────────────────────
$M       = 20       # margen lateral
$TOP     = 92       # y inicio contenido
$LWIDTH  = 340      # ancho panel izquierdo
$GAP     = 14       # separación entre paneles
$RBASE   = $M + $LWIDTH + $GAP                  # x panel scripts = 374
$RWIDTH  = 1280 - $RBASE - $M                   # ancho panel scripts = 886
$CH      = 443      # altura zona contenido
$LOG_Y   = $TOP + $CH + 8                        # = 543
$LOG_H   = 108
$BTN_Y   = $LOG_Y + $LOG_H + 8                  # = 659

# ── Helpers de control ────────────────────────────────────────────────────────
function New-Lbl {
    param([string]$t, [int]$x, [int]$y, [int]$w, [int]$h,
          $font=$fNorm, $fg=$cTxt, $parent=$null)
    $l = New-Object Windows.Forms.Label
    $l.Text=$t; $l.Location="$x,$y"; $l.Size="$w,$h"
    $l.Font=$font; $l.ForeColor=$fg; $l.BackColor=[Drawing.Color]::Transparent
    if ($parent) { $parent.Controls.Add($l) }
    return $l
}

function New-FlatBtn {
    param([string]$t, [int]$x, [int]$y, [int]$w, [int]$h,
          $bg=$cBtnAll, $fg=$cTxt, $font=$fNorm, $parent=$null)
    $b = New-Object Windows.Forms.Button
    $b.Text=$t; $b.Location="$x,$y"; $b.Size="$w,$h"
    $b.BackColor=$bg; $b.ForeColor=$fg; $b.Font=$font
    $b.FlatStyle='Flat'; $b.FlatAppearance.BorderSize=0
    $b.Cursor=[Windows.Forms.Cursors]::Hand
    if ($parent) { $parent.Controls.Add($b) }
    return $b
}

function New-HLine {
    param([int]$y, [int]$w, $parent=$null)
    $p = New-Object Windows.Forms.Panel
    $p.Location="0,$y"; $p.Size="$w,1"; $p.BackColor=$cBorder
    if ($parent) { $parent.Controls.Add($p) }
    return $p
}

# Sección coloreada con cabecera y botón "Todos"
function New-Section {
    param([Windows.Forms.Control]$parent,
          [string]$title, [int]$y, [int]$h, [bool]$showBtn=$true)
    $p = New-Object Windows.Forms.Panel
    $p.Location="0,$y"; $p.Size="$LWIDTH,$h"
    $p.BackColor=$cCard; $p.BorderStyle='FixedSingle'
    $parent.Controls.Add($p)
    New-Lbl $title 10 8 200 20 $fHead $cAccent $p | Out-Null
    New-HLine 30 $LWIDTH $p | Out-Null
    $btn = $null
    if ($showBtn) {
        $btn = New-FlatBtn 'Todos' ($LWIDTH-100) 5 90 20 $cBtnAll $cTxtDim $fSmall $p
    }
    return @{ Panel=$p; BtnAll=$btn }
}

# ── Formulario ────────────────────────────────────────────────────────────────
$form = New-Object Windows.Forms.Form
$form.Text            = 'schoperena-win-setup'
$form.ClientSize      = '1280,720'
$form.StartPosition   = 'CenterScreen'
$form.BackColor       = $cBg
$form.ForeColor       = $cTxt
$form.Font            = $fNorm
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox     = $false

# ── Header ────────────────────────────────────────────────────────────────────
$pHeader = New-Object Windows.Forms.Panel
$pHeader.Location='0,0'; $pHeader.Size='1280,90'; $pHeader.BackColor=$cPanel
$form.Controls.Add($pHeader)

# Logo ASCII "SCH" a la izquierda
$asciiArt = " ___  ___ _  _ `r`n/ __>/ __| || |`r`n\__ \| (__ | >< |`r`n<___/\___|_||_|"
$lblAscii = New-Object Windows.Forms.Label
$lblAscii.Text      = $asciiArt
$lblAscii.Location  = '18,8'
$lblAscii.Size      = '160,64'
$lblAscii.Font      = New-Object Drawing.Font('Consolas', 8, [Drawing.FontStyle]::Bold)
$lblAscii.ForeColor = [Drawing.Color]::FromArgb(0, 140, 168)   # cyan apagado
$lblAscii.BackColor = [Drawing.Color]::Transparent
$pHeader.Controls.Add($lblAscii)

# Título y subtítulo a la derecha del logo
New-Lbl 'WIN SETUP' 184 14 500 42 (New-Object Drawing.Font('Segoe UI',22,[Drawing.FontStyle]::Bold)) $cAccent $pHeader | Out-Null
New-Lbl 'Setup interactivo de entorno personal Windows' 186 58 700 18 $fSmall $cTxtDim $pHeader | Out-Null

New-HLine 90 1280 $form | Out-Null

# ── Panel izquierdo ───────────────────────────────────────────────────────────
$pLeft = New-Object Windows.Forms.Panel
$pLeft.Location="$M,$TOP"; $pLeft.Size="$LWIDTH,$CH"; $pLeft.BackColor=$cBg
$form.Controls.Add($pLeft)

# Navegadores
$secNav = New-Section $pLeft 'Navegadores' 0 152
$clbNav = New-Object Windows.Forms.CheckedListBox
$clbNav.Location='8,34'; $clbNav.Size='324,110'
$clbNav.BackColor=$cCard; $clbNav.ForeColor=$cTxt; $clbNav.Font=$fNorm
$clbNav.BorderStyle='None'; $clbNav.CheckOnClick=$true; $clbNav.IntegralHeight=$false
foreach ($b in $browserItems) { [void]$clbNav.Items.Add($b.label) }
$secNav.Panel.Controls.Add($clbNav)
# Herramientas AI
$secAI = New-Section $pLeft 'Herramientas AI' 162 118
$clbAI = New-Object Windows.Forms.CheckedListBox
$clbAI.Location='8,34'; $clbAI.Size='324,78'
$clbAI.BackColor=$cCard; $clbAI.ForeColor=$cTxt; $clbAI.Font=$fNorm
$clbAI.BorderStyle='None'; $clbAI.CheckOnClick=$true; $clbAI.IntegralHeight=$false
foreach ($a in $aiItems) { [void]$clbAI.Items.Add($a.label) }
$secAI.Panel.Controls.Add($clbAI)
$secAI.BtnAll.Add_Click({
    $a = $clbAI.CheckedItems.Count -eq $clbAI.Items.Count
    for ($i=0; $i -lt $clbAI.Items.Count; $i++) { if (-not $aiInstalledIdx.Contains($i)) { $clbAI.SetItemChecked($i, -not $a) } }
})

# ── Detección de paquetes ya instalados ───────────────────────────────────────
$navInstalledIdx = [System.Collections.Generic.HashSet[int]]::new()
$aiInstalledIdx  = [System.Collections.Generic.HashSet[int]]::new()

# OwnerDraw para CheckedListBox: grises y bloqueados si ya instalado
function Register-CLBOwnerDraw {
    param([Windows.Forms.CheckedListBox]$clb, [System.Collections.Generic.HashSet[int]]$instIdx)
    $clb.DrawMode = [Windows.Forms.DrawMode]::OwnerDrawFixed
    $clb.Add_DrawItem({
        param($s, $e)
        if ($e.Index -lt 0) { return }
        $inst = $instIdx.Contains($e.Index)
        $bg   = if ($inst) { [Drawing.Color]::FromArgb(22,30,22) } else { $cCard }
        $fg   = if ($inst) { [Drawing.Color]::FromArgb(68,105,68) } else { $cTxt }
        $e.Graphics.FillRectangle([Drawing.SolidBrush]::new($bg), $e.Bounds)
        $chkSz = [Windows.Forms.CheckBoxRenderer]::GetGlyphSize($e.Graphics, [Windows.Forms.VisualStyles.CheckBoxState]::CheckedNormal)
        $pt    = New-Object Drawing.Point(($e.Bounds.X+2), ($e.Bounds.Y+[int](($e.Bounds.Height-$chkSz.Height)/2)))
        $st    = if ($s.GetItemChecked($e.Index)) { [Windows.Forms.VisualStyles.CheckBoxState]::CheckedNormal } else { [Windows.Forms.VisualStyles.CheckBoxState]::UncheckedNormal }
        [Windows.Forms.CheckBoxRenderer]::DrawCheckBox($e.Graphics, $pt, $st)
        $tx = [float]($pt.X + $chkSz.Width + 4)
        $ty = [float]($e.Bounds.Y + 2)
        $e.Graphics.DrawString($s.Items[$e.Index], $e.Font, [Drawing.SolidBrush]::new($fg), $tx, $ty)
        if ($inst) {
            $tw = $e.Graphics.MeasureString($s.Items[$e.Index], $e.Font).Width
            $e.Graphics.DrawString('  ✓ ya instalado', $fSmall, [Drawing.SolidBrush]::new([Drawing.Color]::FromArgb(55,125,55)), $tx+$tw, $ty+2)
        }
    }.GetNewClosure())
    $clb.Add_ItemCheck({
        param($s, $e)
        if ($e.NewValue -eq [Windows.Forms.CheckState]::Unchecked -and $instIdx.Contains($e.Index)) {
            $e.NewValue = [Windows.Forms.CheckState]::Checked
        }
    }.GetNewClosure())
}

Register-CLBOwnerDraw $clbNav $navInstalledIdx
Register-CLBOwnerDraw $clbAI  $aiInstalledIdx

# Fix "Todos" para nav también (respetar instalados)
$secNav.BtnAll.Add_Click({
    $a = ($clbNav.CheckedItems.Count - $navInstalledIdx.Count) -eq ($clbNav.Items.Count - $navInstalledIdx.Count)
    for ($i=0; $i -lt $clbNav.Items.Count; $i++) { if (-not $navInstalledIdx.Contains($i)) { $clbNav.SetItemChecked($i, -not $a) } }
})

# Win11Debloat
$secDeb = New-Section $pLeft 'Extras' 290 82 $false
$chkDeb = New-Object Windows.Forms.CheckBox
$chkDeb.Location='10,36'; $chkDeb.Size='316,20'
$chkDeb.Text='Win11Debloat  (by Raphire)'
$chkDeb.BackColor=$cCard; $chkDeb.ForeColor=$cTxt; $chkDeb.Font=$fNorm; $chkDeb.FlatStyle='Flat'
$secDeb.Panel.Controls.Add($chkDeb)
New-Lbl 'Elimina bloatware de Windows 11 — se ejecuta al finalizar el resto' `
    10 58 320 16 $fSmall $cTxtDim $secDeb.Panel | Out-Null

# ── Panel derecho: Scripts ────────────────────────────────────────────────────
$pScripts = New-Object Windows.Forms.Panel
$pScripts.Location="$RBASE,$TOP"; $pScripts.Size="$RWIDTH,$CH"
$pScripts.BackColor=$cCard; $pScripts.BorderStyle='FixedSingle'
$form.Controls.Add($pScripts)

New-Lbl 'Scripts personales' 10 8 300 20 $fHead $cAccent $pScripts | Out-Null
New-HLine 30 $RWIDTH $pScripts | Out-Null

$btnAllScripts = New-FlatBtn 'Todos' ($RWIDTH-100) 5 90 20 $cBtnAll $cTxtDim $fSmall $pScripts

# DataGridView
$dgv = New-Object Windows.Forms.DataGridView
$dgv.Location='5,34'
$dgv.Size=New-Object Drawing.Size(($RWIDTH-10), ($CH-39))
$dgv.BackgroundColor=$cCard; $dgv.GridColor=$cBorder; $dgv.BorderStyle='None'
$dgv.RowHeadersVisible=$false; $dgv.AllowUserToAddRows=$false
$dgv.AllowUserToDeleteRows=$false; $dgv.AllowUserToResizeRows=$false
$dgv.MultiSelect=$false; $dgv.SelectionMode='FullRowSelect'
$dgv.RowTemplate.Height=30; $dgv.ScrollBars='Vertical'
$dgv.EnableHeadersVisualStyles=$false; $dgv.Font=$fNorm

$dgv.DefaultCellStyle.BackColor=$cCard; $dgv.DefaultCellStyle.ForeColor=$cTxt
$dgv.DefaultCellStyle.SelectionBackColor=$cSelBg; $dgv.DefaultCellStyle.SelectionForeColor=$cTxt
$dgv.AlternatingRowsDefaultCellStyle.BackColor=$cRowAlt
$dgv.AlternatingRowsDefaultCellStyle.ForeColor=$cTxt
$dgv.ColumnHeadersDefaultCellStyle.BackColor=$cPanel
$dgv.ColumnHeadersDefaultCellStyle.ForeColor=$cAccent
$dgv.ColumnHeadersDefaultCellStyle.Font=$fHead
$dgv.ColumnHeadersHeight=30; $dgv.ColumnHeadersBorderStyle='Single'

$colChk  = New-Object Windows.Forms.DataGridViewCheckBoxColumn
$colChk.HeaderText=''; $colChk.Width=44; $colChk.Resizable='False'
$colChk.DefaultCellStyle.Alignment='MiddleCenter'
[void]$dgv.Columns.Add($colChk)

$colName = New-Object Windows.Forms.DataGridViewTextBoxColumn
$colName.HeaderText='Script'; $colName.Width=230; $colName.ReadOnly=$true; $colName.Resizable='False'

[void]$dgv.Columns.Add($colName)

$colDesc = New-Object Windows.Forms.DataGridViewTextBoxColumn
$colDesc.HeaderText='Descripción'; $colDesc.AutoSizeMode='Fill'; $colDesc.ReadOnly=$true
[void]$dgv.Columns.Add($colDesc)

$cInstBg = [Drawing.Color]::FromArgb(22, 38, 22)
$cInstFg = [Drawing.Color]::FromArgb(72, 115, 72)

$csDir_ = "$(Split-Path $PROFILE)\CustomScripts"
foreach ($s in $scriptItems) {
    $installed = Test-Path "$csDir_\$($s.name)"
    $desc = if ($installed) { "✓ ya instalado  —  $($s.desc)" } else { $s.desc }
    $idx  = $dgv.Rows.Add($installed, $s.name, $desc)
    if ($installed) {
        $row = $dgv.Rows[$idx]
        $row.Tag = 'installed'
        $row.DefaultCellStyle.BackColor          = $cInstBg
        $row.DefaultCellStyle.ForeColor          = $cInstFg
        $row.DefaultCellStyle.SelectionBackColor = [Drawing.Color]::FromArgb(28,48,28)
        $row.DefaultCellStyle.SelectionForeColor = $cInstFg
    }
}

# Commit inmediato del checkbox + revertir si la fila es 'installed'
$dgv.Add_CurrentCellDirtyStateChanged({
    if ($dgv.IsCurrentCellDirty) {
        $dgv.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})
$dgv.Add_CellValueChanged({
    param($sender, $e)
    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0 -and $dgv.Rows[$e.RowIndex].Tag -eq 'installed') {
        $dgv.Rows[$e.RowIndex].Cells[0].Value = $true
    }
})

$pScripts.Controls.Add($dgv)

$btnAllScripts.Add_Click({
    $any = ($dgv.Rows | Where-Object { $_.Tag -ne 'installed' -and -not [bool]$_.Cells[0].Value }).Count -gt 0
    foreach ($row in $dgv.Rows) { if ($row.Tag -ne 'installed') { $row.Cells[0].Value = $any } }
    $dgv.RefreshEdit()
})

# ── Log area ──────────────────────────────────────────────────────────────────
$rtbLog = New-Object Windows.Forms.RichTextBox
$rtbLog.Location="$M,$LOG_Y"
$rtbLog.Size=New-Object Drawing.Size((1280-$M*2), $LOG_H)
$rtbLog.BackColor=$cPanel; $rtbLog.ForeColor=$cTxtDim; $rtbLog.Font=$fMono
$rtbLog.ReadOnly=$true; $rtbLog.BorderStyle='None'; $rtbLog.ScrollBars='Vertical'
$rtbLog.Text="  Configura tu selección y presiona  ▶ Iniciar Setup  para comenzar.`n"
$form.Controls.Add($rtbLog)

# ── Footer ────────────────────────────────────────────────────────────────────
$btnCancel = New-FlatBtn 'Cancelar' (1280-$M-220) $BTN_Y 100 34 $cBtnCan ([Drawing.Color]::FromArgb(255,120,120)) $fNorm $form
$btnStart  = New-FlatBtn '▶  Iniciar Setup' (1280-$M-110) $BTN_Y 110 34 $cBtnRun $cTxt $fHead $form
$btnCancel.Add_Click({ $form.Close() })

# ── Colormap para el log ──────────────────────────────────────────────────────
$logColors = @{
    green  = $cGreen; cyan   = $cAccent; yellow = $cYellow
    red    = $cRed;   dim    = $cTxtDim; default = $cTxt
}

$appendLog = {
    param([string]$raw)
    if ($raw -match '^\[(\w+)\](.*)') {
        $col  = if ($logColors.ContainsKey($matches[1])) { $logColors[$matches[1]] } else { $cTxt }
        $text = $matches[2]
    } else { $col = $cTxt; $text = $raw }
    $rtbLog.SelectionStart  = $rtbLog.TextLength
    $rtbLog.SelectionLength = 0
    $rtbLog.SelectionColor  = $col
    $rtbLog.AppendText("$text`n")
    $rtbLog.ScrollToCaret()
}

# ── Lógica de inicio ──────────────────────────────────────────────────────────
$btnStart.Add_Click({
    # Recopilar selecciones desde la UI
    $selBrowsers = @(); for ($i=0; $i -lt $clbNav.Items.Count; $i++) { if ($clbNav.GetItemChecked($i)) { $selBrowsers += $browserItems[$i] } }
    $selAI       = @(); for ($i=0; $i -lt $clbAI.Items.Count;  $i++) { if ($clbAI.GetItemChecked($i))  { $selAI       += $aiItems[$i]      } }
    $selScripts  = @()
    foreach ($row in $dgv.Rows) {
        if ([bool]$row.Cells[0].Value) {
            $n = $row.Cells[1].Value
            $selScripts += ($scriptItems | Where-Object { $_.name -eq $n })
        }
    }
    $runDebloat = $chkDeb.Checked

    $btnStart.Enabled = $false; $btnCancel.Enabled = $false
    $rtbLog.Clear()

    # Hashtable sincronizado para comunicación inter-runspace
    $sync = [hashtable]::Synchronized(@{
        Log  = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        Done = $false
    })

    $psDir   = Split-Path $PROFILE
    $csDir   = "$psDir\CustomScripts"
    $modDir  = "$psDir\Modules"
    $profile_ = $PROFILE

    # Script de instalación (corre en runspace separado)
    $installBlock = {
        param($sync, $repoBase, $selBrowsers, $selAI, $selScripts,
              $runDebloat, $psDir, $csDir, $modDir, $profilePath)

        function Log {
            param([string]$msg, [string]$c='default')
            $sync.Log.Enqueue("[$c]$msg")
        }

        function Install-Pkg {
            param([string]$Id, [string]$Name, [string]$Type, [string]$Src='winget')
            try {
                if ($Type -eq 'winget') {
                    $already = winget list --id $Id --accept-source-agreements 2>&1 | Select-String $Id
                    if ($already) { Log "  --  $Name ya instalado" 'dim'; return }
                    Log "  >>  Instalando $Name..." 'cyan'
                    $out = winget install --id $Id -e --accept-package-agreements --accept-source-agreements --source $Src 2>&1
                    foreach ($line in $out) { $s = "$line".Trim(); if ($s) { Log "      $s" 'dim' } }
                } elseif ($Type -eq 'npm') {
                    if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
                        Log "  !!  npm no disponible — instala Node.js primero" 'yellow'; return
                    }
                    $already = npm list -g --depth=0 2>$null | Select-String ([regex]::Escape($Id))
                    if ($already) { Log "  --  $Name ya instalado" 'dim'; return }
                    Log "  >>  Instalando $Name (npm)..." 'cyan'
                    $out = npm install -g $Id 2>&1
                    foreach ($line in $out) { $s = "$line".Trim(); if ($s) { Log "      $s" 'dim' } }
                }
                Log "  OK  $Name" 'green'
            } catch { Log "  !!  $Name : $($_.Exception.Message)" 'yellow' }
        }

        function Deploy {
            param([string]$src, [string]$dst)
            try {
                New-Item -ItemType Directory (Split-Path $dst) -Force | Out-Null
                Copy-Item $src $dst -Force
                Log "  OK  $(Split-Path $dst -Leaf)" 'green'
            } catch { Log "  !!  $(Split-Path $dst -Leaf): $($_.Exception.Message)" 'yellow' }
        }

        try {
            # ── 1. Perfil y archivos de configuración PRIMERO ─────────────────
            # Esto evita que procesos hijo (npm post-install, winget scripts)
            # carguen un perfil antiguo con llamadas a apps no instaladas aún.
            Log '══ Perfil y configuración' 'cyan'
            New-Item -ItemType Directory $psDir,$modDir,$csDir,"$env:APPDATA\fastfetch" -Force | Out-Null
            if (Test-Path $profilePath) {
                Copy-Item $profilePath "$profilePath.bak_$(Get-Date -Format 'yyyyMMdd_HHmmss')" -Force
                Log "  --  Backup del perfil anterior creado" 'dim'
            }
            @(
                @("$repoBase\powershell\profile.ps1",                $profilePath)
                @("$repoBase\powershell\powershell.config.json",     "$psDir\powershell.config.json")
                @("$repoBase\powershell\themes\night-owl.omp.json",  "$psDir\night-owl.omp.json")
                @("$repoBase\powershell\themes\quick-term.omp.json", "$psDir\quick-term.omp.json")
                @("$repoBase\powershell\themes\mytheme.omp.json",    "$psDir\.mytheme.omp.json")
                @("$repoBase\fastfetch\config.jsonc",                "$env:APPDATA\fastfetch\config.jsonc")
            ) | ForEach-Object { Deploy $_[0] $_[1] }

            if ($selScripts.Count -gt 0) {
                Log '══ Scripts personales' 'cyan'
                foreach ($s in $selScripts) { Deploy "$repoBase\scripts\$($s.name)" "$csDir\$($s.name)" }
            }

            $imgConvSrc = "$repoBase\modules\ImgConv"
            if (Test-Path $imgConvSrc) {
                Log '══ Módulo ImgConv' 'cyan'
                $dst = "$modDir\ImgConv"
                New-Item -ItemType Directory $dst -Force | Out-Null
                Get-ChildItem $imgConvSrc | ForEach-Object { Deploy $_.FullName "$dst\$($_.Name)" }
            }

            # ── 2. Paquetes ───────────────────────────────────────────────────
            Log '══ PowerShell 7' 'cyan'
            Install-Pkg 'Microsoft.PowerShell' 'PowerShell 7' 'winget'

            Log '══ Dependencias' 'cyan'
            Install-Pkg 'Git.Git'                  'Git'         'winget'
            Install-Pkg 'GitHub.cli'               'GitHub CLI'  'winget'
            Install-Pkg 'JanDeDobbeleer.OhMyPosh' 'oh-my-posh' 'winget'
            Install-Pkg 'ImageMagick.Q16-HDRI'    'ImageMagick' 'winget'
            Install-Pkg 'Fastfetch-cli.Fastfetch' 'fastfetch'   'winget'
            if ($selAI | Where-Object { $_.type -eq 'npm' }) {
                Install-Pkg 'OpenJS.NodeJS.LTS' 'Node.js LTS' 'winget'
            }

            Log '══ Apps esenciales' 'cyan'
            Install-Pkg 'VideoLAN.VLC'               'VLC'     'winget'
            Install-Pkg 'Microsoft.VisualStudioCode' 'VS Code' 'winget'
            Install-Pkg 'M2Team.NanaZip'             'NanaZip' 'winget'

            if ($selBrowsers.Count -gt 0) {
                Log '══ Navegadores' 'cyan'
                foreach ($b in $selBrowsers) { Install-Pkg $b.id $b.label 'winget' }
            }
            if ($selAI.Count -gt 0) {
                Log '══ Herramientas AI' 'cyan'
                foreach ($a in $selAI) { Install-Pkg $a.id $a.label $a.type }
            }

            Log '══ Módulos PSGallery' 'cyan'
            foreach ($mod in @('Terminal-Icons','ps2exe')) {
                if (Get-Module -ListAvailable -Name $mod) { Log "  --  $mod ya instalado" 'dim'; continue }
                Install-Module -Name $mod -Scope CurrentUser -Force -SkipPublisherCheck
                Log "  OK  $mod" 'green'
            }

            Log '══ FiraCode Nerd Font' 'cyan'
            $fontDst = "$env:LOCALAPPDATA\Microsoft\Windows\Fonts\FiraCodeNerdFont-Regular.ttf"
            if (Test-Path $fontDst) {
                Log "  --  FiraCode Nerd Font ya instalada" 'dim'
            } else {
                $zip = "$env:TEMP\FC_NF.zip"; $dir = "$env:TEMP\FC_NF"
                Invoke-WebRequest 'https://github.com/ryanoasis/nerd-fonts/releases/latest/download/FiraCode.zip' `
                    -OutFile $zip -UseBasicParsing
                Expand-Archive $zip $dir -Force
                $ttf = Get-ChildItem "$dir\*.ttf" | Where-Object { $_.Name -like '*Regular*' } | Select-Object -First 1
                if ($ttf) {
                    New-Item -ItemType Directory "$env:LOCALAPPDATA\Microsoft\Windows\Fonts" -Force | Out-Null
                    Copy-Item $ttf.FullName $fontDst -Force
                    New-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' `
                        -Name 'FiraCode Nerd Font Regular (TrueType)' -Value $fontDst -PropertyType String -Force | Out-Null
                    Log "  OK  FiraCode Nerd Font Regular" 'green'
                }
                Remove-Item $zip,$dir -Recurse -Force -ErrorAction SilentlyContinue
            }

            if ($runDebloat) {
                Log '══ Win11Debloat (Raphire) — iniciando...' 'cyan'
                & ([scriptblock]::Create((irm 'https://debloat.raphi.re/')))
            }

            Log '' 'default'
            Log '✅  Setup completado. Reinicia Windows Terminal.' 'green'
        } catch {
            Log "❌  Error inesperado: $($_.Exception.Message)" 'red'
        }
        $sync.Done = $true
    }

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()
    $ps = [powershell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript($installBlock)
    [void]$ps.AddArgument($sync)
    [void]$ps.AddArgument($RepoBase)
    [void]$ps.AddArgument($selBrowsers)
    [void]$ps.AddArgument($selAI)
    [void]$ps.AddArgument($selScripts)
    [void]$ps.AddArgument($runDebloat)
    [void]$ps.AddArgument($psDir)
    [void]$ps.AddArgument($csDir)
    [void]$ps.AddArgument($modDir)
    [void]$ps.AddArgument($profile_)
    $handle = $ps.BeginInvoke()

    # Timer: bombea el log del runspace a la UI sin bloquear
    $timer = New-Object Windows.Forms.Timer; $timer.Interval=120
    $timer.Add_Tick({
        [string]$item = $null
        while ($sync.Log.TryDequeue([ref]$item)) { & $appendLog $item }
        if ($sync.Done) {
            $timer.Stop()
            try { $ps.EndInvoke($handle) } catch {}
            $rs.Close()
            $btnStart.Text='✅  Completado'; $btnStart.BackColor=$cGreen
            $btnCancel.Enabled=$true; $btnCancel.Text='Cerrar'
        }
    }.GetNewClosure())
    $timer.Start()
})

# ── Detección síncrona de paquetes instalados (antes de mostrar el form) ─────
Write-Host '  Detectando paquetes instalados...' -ForegroundColor DarkCyan
$_wl = ''; $_nl = ''
try { $_wl = (winget list --accept-source-agreements 2>&1) -join "`n" } catch {}
try {
    if (Get-Command npm -ErrorAction SilentlyContinue) {
        $_nl = (npm list -g --depth=0 2>$null) -join "`n"
    }
} catch {}

for ($i = 0; $i -lt $browserItems.Count; $i++) {
    if ($_wl -match [regex]::Escape($browserItems[$i].id)) {
        [void]$navInstalledIdx.Add($i); $clbNav.SetItemChecked($i, $true)
    }
}
for ($i = 0; $i -lt $aiItems.Count; $i++) {
    $found = if ($aiItems[$i].type -eq 'winget') { $_wl -match [regex]::Escape($aiItems[$i].id) }
             else { $_nl -match [regex]::Escape($aiItems[$i].id) }
    if ($found) { [void]$aiInstalledIdx.Add($i); $clbAI.SetItemChecked($i, $true) }
}

# ── Mostrar formulario ────────────────────────────────────────────────────────
[void]$form.ShowDialog()
