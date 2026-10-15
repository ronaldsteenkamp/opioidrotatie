# Installeert de rekenhulp Opioïdrotatie voor de huidige gebruiker:
#  - kopieert de werkmap en de twee uitleesscripts naar een lokale map
#    (standaard Downloads\Opioidrotatie; vanaf een netwerkschijf blokkeert Excel vaak de macro's);
#  - maakt op het bureaublad de snelkoppelingen "Opioidrotatie morfinepomp" (werkmap) en
#    "Lees opioiden uit Pharmacom" met sneltoets Ctrl+Alt+F3;
#  - controleert of de Java Access Bridge (koppeling met Pharmacom) aanstaat en zet hem zo nodig
#    aan (jabswitch -enable uit de Java-map van Pharmacom, anders .accessibility.properties zelf).
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

# Java Access Bridge: nodig voor het uitlezen uit Pharmacom. Java leest bij het starten
# %USERPROFILE%\.accessibility.properties; jabswitch -enable schrijft dat bestand.
$JabBestand = Join-Path $env:USERPROFILE '.accessibility.properties'
function Test-JabAan {
    (Test-Path -LiteralPath $JabBestand) -and
    [bool](Select-String -LiteralPath $JabBestand -Pattern '^\s*assistive_technologies\s*=.*AccessBridge' -Quiet)
}
$Pharmacom = Get-Process javaw -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like 'Pharmacom*' } | Select-Object -First 1
$JreBin = if ($Pharmacom) { Split-Path $Pharmacom.Path } else {
    # Pharmacom niet open: nieuwste Java van PharmaPartners
    Get-ChildItem 'C:\PharmaPartners\jres\*\*\bin\jabswitch.exe' -ErrorAction SilentlyContinue |
        Sort-Object { $_.VersionInfo.FileVersionRaw } | Select-Object -Last 1 | ForEach-Object { $_.DirectoryName }
}
$JabWasAan = Test-JabAan
if (-not $JabWasAan) {
    $Jabswitch = if ($JreBin) { Join-Path $JreBin 'jabswitch.exe' }
    if ($Jabswitch -and (Test-Path -LiteralPath $Jabswitch)) {
        try { Start-Process -FilePath $Jabswitch -ArgumentList '-enable' -WindowStyle Hidden -Wait } catch {}
    }
    if (-not (Test-JabAan)) {
        # Zonder jabswitch: het instellingenbestand zelf schrijven (zoals jabswitch doet)
        try { Add-Content -LiteralPath $JabBestand -Encoding ASCII -Value "`nassistive_technologies=com.sun.java.accessibility.AccessBridge`nscreen_magnifier_present=true" } catch {}
    }
}
if (-not (Test-JabAan)) {
    $Jab = "LET OP: de Java Access Bridge kon niet ingeschakeld worden. Voer handmatig uit:`n" +
           $(if ($JreBin) { "`"$JreBin\jabswitch.exe`" -enable" } else { 'jabswitch -enable (uit de Java-map van Pharmacom)' }) +
           "`nen start Pharmacom daarna opnieuw."
}
elseif ($Pharmacom) {
    # Draait Pharmacom al met de bridge? Alleen 32-bit PowerShell ziet de modules van 32-bit Java.
    $Ps32 = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    $Geladen = try { & $Ps32 -NoProfile -Command "[bool]((Get-Process -Id $($Pharmacom.Id) -EA 0).Modules | ? ModuleName -like 'JavaAccessBridge*')" 2>$null } catch { '' }
    $Jab = if ($Geladen -eq 'True') { 'De koppeling met Pharmacom (Java Access Bridge) staat aan.' }
           else { "De koppeling met Pharmacom (Java Access Bridge) $(if ($JabWasAan) { 'staat aan' } else { 'is nu ingeschakeld' }), maar het open Pharmacom gebruikt hem nog niet: sluit Pharmacom en start het opnieuw." }
}
elseif ($JabWasAan) { $Jab = 'De koppeling met Pharmacom (Java Access Bridge) staat aan.' }
else { $Jab = 'De koppeling met Pharmacom (Java Access Bridge) is nu ingeschakeld; hij werkt zodra Pharmacom (opnieuw) gestart wordt.' }

$Tekst = "Geinstalleerd in:`n$Map`n`nOp het bureaublad staan nu 'Opioidrotatie morfinepomp' en 'Lees opioiden uit Pharmacom' (sneltoets Ctrl+Alt+F3).`n`n" +
         "$Jab`n`nKlik bij het eerste openen van de werkmap op 'Inhoud inschakelen'."
if ($Stil) { $Tekst } else {
    Add-Type -AssemblyName System.Windows.Forms
    [void][Windows.Forms.MessageBox]::Show($Tekst, 'Opioidrotatie', 'OK', 'Information')
}
