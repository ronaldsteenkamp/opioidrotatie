# Leest de opioiden uit het open Pharmacom-dossier (medicatiestatus, F3) en zet ze in de
# rekenhulp Opioidrotatie_morfinepomp.xlsm (blok 1 op tabblad Rotatie).
#
# Wordt gestart via de snelkoppeling/sneltoets, NIET vanuit Excel: op veel pc's mogen macro's
# in Office geen PowerShell starten (beveiliging). Dit script start zelf het uitleesscript
# en geeft het resultaat via COM aan de macro Verwerk in de werkmap.
# Het script haalt zelf Pharmacom naar voren en drukt zo nodig F3 (medicatiestatus). Daarna haalt het de
# nierfunctie (eGFR) uit het Diagnostisch dossier (Alt+I, Alt+D) en gaat terug naar de medicatiestatus (F3).
# Toetsen gaan alleen naar Pharmacom als dat het actieve venster is en er geen melding openstaat.
# Daarna komt Excel naar voren.
#
# Starten met 32-bit PowerShell:
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File Pharmacom_naar_Excel.ps1
# -Proef: koppeling wel starten, maar verzonnen uitvoer in plaats van het dossier (testen).
# -Stil:  geen meldingen in Excel (testen).
param([switch]$Proef, [switch]$Stil)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$Naam = 'Opioidrotatie_morfinepomp.xlsm'
$Werkmap = Join-Path $PSScriptRoot $Naam
$Lezer = Join-Path $PSScriptRoot 'Pharmacom_medicatiestatus.ps1'

Add-Type @"
using System; using System.Runtime.InteropServices;
public static class Venster {
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr h);
  public static uint VoorgrondPid() { uint p; GetWindowThreadProcessId(GetForegroundWindow(), out p); return p; }
}
"@

function Meld($Tekst, $Icoon = 'Warning') {
    if (-not $Stil) { [void][Windows.Forms.MessageBox]::Show($Tekst, 'Lees uit Pharmacom', 'OK', $Icoon) }
    else { $Tekst }
}

function Pharmacom { Get-Process javaw -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like 'Pharmacom*' } | Select-Object -First 1 }

# Stuurt toetsen naar Pharmacom, alleen als het het actieve venster is en er geen melding openstaat
function ToetsenNaarPharmacom([string[]]$Toetsen) {
    $P = Pharmacom
    if (-not $P -or -not [Venster]::IsWindowEnabled($P.MainWindowHandle)) { return $false }
    if ([Venster]::VoorgrondPid() -ne $P.Id) { [void](ActiveerPharmacom $P) }
    foreach ($K in $Toetsen) {
        if ([Venster]::VoorgrondPid() -ne $P.Id) { return $false }
        [System.Windows.Forms.SendKeys]::SendWait($K)
        Start-Sleep -Milliseconds 600
    }
    return $true
}

# Haalt Pharmacom naar voren; true als het daarna echt het actieve venster is
function ActiveerPharmacom($P) {
    $H = $P.MainWindowHandle
    if ([Venster]::IsIconic($H)) { [void][Venster]::ShowWindow($H, 9) }   # herstellen
    try { [void](New-Object -ComObject WScript.Shell).AppActivate($P.Id) } catch { }
    [void][Venster]::SetForegroundWindow($H)
    $Eind = (Get-Date).AddSeconds(2)
    while ((Get-Date) -lt $Eind) {
        if ([Venster]::VoorgrondPid() -eq $P.Id) { return $true }
        Start-Sleep -Milliseconds 50
    }
    return $false
}

# Leest de medicatiestatus, of met -Lab de nierfunctie (via een tijdelijk bestand dat direct weer weg gaat)
function LeesUit([switch]$Lab) {
    $Tmp = Join-Path $env:TEMP ("opioidrotatie_" + [guid]::NewGuid().ToString('N') + ".txt")
    $Arg = @{ Uit = $Tmp }
    if ($Proef) { $Arg.Proef = $true }
    if ($Lab) { $Arg.Lab = $true }
    & $Lezer @Arg
    if (-not (Test-Path -LiteralPath $Tmp)) { return "" }
    $T = [IO.File]::ReadAllText($Tmp, [Text.Encoding]::UTF8)
    [IO.File]::Delete($Tmp)
    return $T
}

try {
    # 1. Pharmacom naar voren halen en controleren dat er een dossier open is
    $P = Pharmacom
    if (-not $P) { Meld "Pharmacom is niet geopend."; exit 1 }
    $Actief = ActiveerPharmacom $P
    if (-not $Proef -and $P.MainWindowTitle -notlike '*Dossier*') {
        Meld "Open eerst het dossier van de patiënt in Pharmacom (venster: $($P.MainWindowTitle))."; exit 1
    }

    # 2. Uitlezen; staat het dossier niet op de medicatiestatus, dan F3 en opnieuw.
    #    F3 wordt alleen gestuurd als Pharmacom het actieve venster is en er geen
    #    melding openstaat (dan is het hoofdvenster uitgeschakeld).
    $Tekst = LeesUit
    $Poging = 0
    while ($Tekst -like 'FOUT*niet gevonden*' -and $Poging -lt 2) {
        $Poging++
        $P = Pharmacom
        if (-not $P -or -not [Venster]::IsWindowEnabled($P.MainWindowHandle)) { break }
        if ([Venster]::VoorgrondPid() -ne $P.Id) { $Actief = ActiveerPharmacom $P }
        if ([Venster]::VoorgrondPid() -ne $P.Id) { break }
        [System.Windows.Forms.SendKeys]::SendWait('{F3}')
        Start-Sleep -Milliseconds (1200 * $Poging)
        $Tekst = LeesUit
    }
    if ($Tekst -eq "") { Meld "Het uitleesscript gaf geen antwoord."; exit 1 }

    # Nierfunctie: Diagnostisch dossier (Alt+I, Alt+D), daarna terug naar de medicatiestatus (F3).
    # De eGFR komt achter de venstertitel op de eerste regel: OK<tab>titel<tab>eGFR<tab>datum<tab>naam.
    if ($Tekst -like 'OK*') {
        $Nier = LeesUit -Lab
        if ($Nier -notlike 'OK*' -and -not $Proef) {
            if (ToetsenNaarPharmacom @('%i', '%d')) {
                Start-Sleep -Milliseconds 800
                $Nier = LeesUit -Lab
                [void](ToetsenNaarPharmacom @('{F3}'))
            }
        }
        $Eerste, $Rest = $Tekst -split "`n", 2
        $Velden = $Nier.Trim() -split "`t"
        # Alleen overnemen als het hetzelfde dossier is
        if ($Nier -like 'OK*' -and $Velden.Count -ge 5 -and $Velden[1] -eq ($Eerste -split "`t")[1]) {
            $Tekst = $Eerste.TrimEnd() + "`t" + $Velden[2] + "`t" + $Velden[3] + "`t" + $Velden[4] + "`n" + $Rest
        }
    }

    # 2. Excel: de open werkmap gebruiken, anders openen
    try { $Xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $Xl = $null }
    $Wb = $null; $NieuweExcel = $false; $Geopend = $false
    if ($Xl) { foreach ($W in $Xl.Workbooks) { if ($W.Name -eq $Naam) { $Wb = $W; break } } }
    if (-not $Wb) {
        if (-not (Test-Path -LiteralPath $Werkmap)) { Meld "De rekenhulp staat niet naast dit script:`n$Werkmap"; exit 1 }
        if (-not $Xl) { $Xl = New-Object -ComObject Excel.Application; $NieuweExcel = $true }
        $Wb = $Xl.Workbooks.Open($Werkmap)
        $Geopend = $true
    }
    if (-not $Stil) {
        $Xl.Visible = $true
        $Wb.Activate()
        $Xl.WindowState = -4137   # gemaximaliseerd
        try { [void](New-Object -ComObject WScript.Shell).AppActivate($Xl.Caption) } catch { }
    }

    # 3. Verwerken in de werkmap (vult blok 1 en toont de samenvatting)
    $Res = $Xl.Run("'$Naam'!Verwerk", $Tekst, [bool]$Stil)
    if ($Stil) {
        # Testen: resultaat tonen en wat dit script opende weer sluiten (zonder opslaan)
        $Res
        $Ws = $Wb.Worksheets.Item('Rotatie')
        "rij 10: " + $Ws.Range('B10').Text + " | " + $Ws.Range('C10').Text + " | " + $Ws.Range('D10').Text + " | " + $Ws.Range('I10').Text
        if ($Geopend) { $Wb.Close($false) }
        if ($NieuweExcel) { $Xl.Quit() }
    }
} catch {
    Meld ("Uitlezen lukte niet: " + $_.Exception.Message) 'Error'
    exit 1
}
