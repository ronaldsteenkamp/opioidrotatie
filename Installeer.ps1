# Installeert de rekenhulp Opioïdrotatie voor de huidige gebruiker:
#  - kopieert de werkmap en de twee uitleesscripts naar een lokale map
#    (standaard Downloads\Opioidrotatie; vanaf een netwerkschijf blokkeert Excel vaak de macro's);
#  - maakt op het bureaublad de snelkoppelingen "Opioidrotatie morfinepomp" (werkmap) en
#    "Lees opioiden uit Pharmacom" met sneltoets Ctrl+Alt+F3.
#
# Gebruik (in de uitgepakte map):
#   powershell -ExecutionPolicy Bypass -File Installeer.ps1 [-Map <doelmap>] [-Stil]
param([string]$Map = "", [switch]$Stil)
$ErrorActionPreference = 'Stop'

$Bestanden = 'Opioidrotatie_morfinepomp.xlsm', 'Pharmacom_naar_Excel.ps1', 'Pharmacom_medicatiestatus.ps1'
foreach ($B in $Bestanden) {
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $B))) {
        throw "$B ontbreekt naast Installeer.ps1. Download de zip van de nieuwste release (of bouw eerst met bron\Bouw_werkmap.ps1)."
    }
}

if (-not $Map) {
    $Downloads = [Environment]::ExpandEnvironmentVariables((Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders').'{374DE290-123F-4565-9164-39C4925E467B}')
    $Map = Join-Path $Downloads 'Opioidrotatie'
}
New-Item -ItemType Directory -Force $Map | Out-Null
foreach ($B in $Bestanden) { Copy-Item -LiteralPath (Join-Path $PSScriptRoot $B) -Destination $Map -Force }

# Snelkoppelingen op het bureaublad
$Sh = New-Object -ComObject WScript.Shell
$Bureaublad = [Environment]::GetFolderPath('Desktop')
$Excel = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue).'(default)'

$W = $Sh.CreateShortcut((Join-Path $Bureaublad 'Opioidrotatie morfinepomp.lnk'))
$W.TargetPath = Join-Path $Map 'Opioidrotatie_morfinepomp.xlsm'
$W.WorkingDirectory = $Map
$W.Description = 'Rekenhulp opioidrotatie naar morfine s.c./i.v.-pomp'
$W.Save()

$L = $Sh.CreateShortcut((Join-Path $Bureaublad 'Lees opioiden uit Pharmacom.lnk'))
$L.TargetPath = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"   # 32-bit: Pharmacom draait op 32-bit Java
$L.Arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$(Join-Path $Map 'Pharmacom_naar_Excel.ps1')`""
$L.WorkingDirectory = $Map
$L.WindowStyle = 7
$L.Hotkey = 'CTRL+ALT+F3'
if ($Excel) { $L.IconLocation = "$Excel,0" }
$L.Description = 'Opioiden uit het open Pharmacom-dossier in de rekenhulp zetten'
$L.Save()

$Tekst = "Geinstalleerd in:`n$Map`n`nOp het bureaublad staan nu 'Opioidrotatie morfinepomp' en 'Lees opioiden uit Pharmacom' (sneltoets Ctrl+Alt+F3).`n`n" +
         "Klik bij het eerste openen van de werkmap op 'Inhoud inschakelen'. Voor het uitlezen uit Pharmacom moet de Java Access Bridge aanstaan (zie README)."
if ($Stil) { $Tekst } else {
    Add-Type -AssemblyName System.Windows.Forms
    [void][Windows.Forms.MessageBox]::Show($Tekst, 'Opioidrotatie', 'OK', 'Information')
}
