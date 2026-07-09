# schoperena-win-setup

Configuración personal de PowerShell: prompt, módulos, temas y scripts.

## Setup en un equipo nuevo

No necesitas `git`. Abre **Windows PowerShell** (el que viene con Windows 11) y ejecuta:

**Consola (compatible con PS 5.1 y PS 7):**
```powershell
& ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/schoperena/schoperena-win-setup/main/setup.ps1')))
```

**GUI 1280×720 (requiere PS 7):**
```powershell
& ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/schoperena/schoperena-win-setup/main/setup-gui.ps1')))
```

El script de consola detecta automáticamente si estás en PowerShell 5.1, instala PowerShell 7 con `winget` y se relanza en él. La versión GUI asume PS 7 ya instalado.

Si prefieres clonar el repo primero:

```powershell
git clone https://github.com/schoperena/schoperena-win-setup "$env:USERPROFILE\.win-setup"
& "$env:USERPROFILE\.win-setup\setup.ps1"
```

## ¿Qué instala?

El script pide selección interactiva para navegadores, herramientas AI y scripts personales. El resto se instala siempre.

| Componente | Origen | Selección |
|---|---|---|
| **PowerShell 7** | winget | siempre |
| **Git** | winget | siempre |
| **GitHub CLI** | winget | siempre |
| **oh-my-posh** | winget | siempre |
| **ImageMagick** | winget | siempre |
| **fastfetch** | winget | siempre |
| **FiraCode Nerd Font** | GitHub releases | siempre |
| **VLC** | winget | siempre |
| **Visual Studio Code** | winget | siempre |
| **NanaZip** | winget | siempre |
| **WhatsApp** | Microsoft Store | siempre |
| **Chrome / Brave / Firefox / LibreWolf** | winget | multi-selección |
| **Claude Desktop** | winget | multi-selección |
| **Claude Code** | npm | multi-selección |
| **Codex CLI** | npm | multi-selección |
| **Terminal-Icons** | PSGallery | siempre |
| **ps2exe** | PSGallery | siempre |
| **ImgConv** | este repo | siempre |
| **Temas OMP** | este repo | siempre |
| **Scripts personales** | este repo | multi-selección |

También configura automáticamente **Windows Terminal**: FiraCode Nerd Font en todos los perfiles y PowerShell 7 como perfil por defecto.

Al final del setup se ofrece ejecutar **[Win11Debloat](https://github.com/raphire/win11debloat)** de [Raphire](https://github.com/raphire) — una herramienta interactiva para eliminar bloatware de Windows 11. Créditos completos al autor original.

## Estructura del repo y destinos de instalación

```
schoperena-win-setup/                      destino en el equipo
├── powershell/
│   ├── profile.ps1              →  $PROFILE
│   ├── powershell.config.json   →  ~\Documents\PowerShell\
│   └── themes/
│       ├── night-owl.omp.json   →  ~\Documents\PowerShell\          ← tema activo
│       ├── quick-term.omp.json  →  ~\Documents\PowerShell\
│       └── mytheme.omp.json     →  ~\Documents\PowerShell\.mytheme.omp.json
├── fastfetch/
│   └── config.jsonc             →  %APPDATA%\fastfetch\config.jsonc
├── scripts/                     →  ~\Documents\PowerShell\CustomScripts\ (selección)
│   ├── MenuScripts.ps1                    (toolbox — hub de scripts)
│   ├── RenombrarMasivo.ps1
│   ├── BloquearAdobe.ps1
│   ├── New-SSHKey.ps1
│   ├── New-QRCode.ps1
│   ├── DividirPDF.ps1
│   ├── FormatearDisco.ps1
│   ├── deblotear_TCL10L.ps1
│   ├── stirling-sch.ps1
│   ├── tree.ps1
│   ├── verify-checksum.ps1
│   ├── win11_rpd_patch.ps1
│   ├── calc_digito_de_verificacion.py
│   └── procesar_notebook.py
├── Modules/
│   └── ImgConv/                 →  ~\Documents\PowerShell\Modules\ImgConv\
├── setup.ps1
└── README.md
```

## Comandos rápidos

| Comando | Descripción |
|---|---|
| `toolbox` | Abre el hub de scripts personales |
| `ImgConv` | Convierte imágenes (HEIC, PNG, JPG, etc.) |

## Scripts en `toolbox`

| Script | Descripción |
|---|---|
| `MenuScripts.ps1` | HUB central para lanzar todos los scripts del toolbox |
| `RenombrarMasivo.ps1` | Renombrado masivo con criterios múltiples (prefijo, sufijo, fecha, numeración) y opción de revertir |
| `BloquearAdobe.ps1` | Bloquea Adobe vía archivo hosts para evitar conexiones no deseadas |
| `New-SSHKey.ps1` | Genera clave SSH (Ed25519 o RSA 4096) para GitHub/GitLab |
| `New-QRCode.ps1` | Genera códigos QR (enlace, WiFi, vCard, email, SMS, geo…) en PNG/SVG con colores a elección |
| `DividirPDF.ps1` | Divide un cartel/plano PDF en hojas A4/A3 imprimibles, con solape y marcas de corte (instala PyMuPDF solo) |
| `FormatearDisco.ps1` | Formatea discos externos (NTFS / exFAT / FAT32) — requiere Admin |
| `deblotear_TCL10L.ps1` | Elimina bloatware del TCL 10L vía ADB |
| `stirling-sch.ps1` | Instala Stirling-PDF apuntando al servidor interno |
| `tree.ps1` | Muestra árbol de directorios |
| `verify-checksum.ps1` | Verifica checksum SHA256/SHA1/MD5 de un archivo |
| `win11_rpd_patch.ps1` | Parche para habilitar RDP en Windows 11 Home |
| `calc_digito_de_verificacion.py` | Calcula dígito de verificación para NIT (Colombia) |
| `procesar_notebook.py` | Convierte y procesa Jupyter Notebooks a distintos formatos |
