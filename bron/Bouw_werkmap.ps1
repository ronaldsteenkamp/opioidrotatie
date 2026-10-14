# Bouwt Opioidrotatie_morfinepomp.xlsm (in de hoofdmap van het project):
# werkmap (werkmap_opbouw.py, openpyxl) + VBA-module Pharmacom.bas + knoppen.
#
# Nodig: Windows, Excel (met "Toegang tot het objectmodel van het VBA-project vertrouwen"),
#        Python 3 met openpyxl.
# Gebruik: powershell -ExecutionPolicy Bypass -File bron\Bouw_werkmap.ps1 [-Installeer]
#   -Installeer: daarna ook Installeer.ps1 draaien (kopie naar Downloads\Opioidrotatie + snelkoppelingen).
param([switch]$Installeer)
$ErrorActionPreference = 'Stop'
$Bron = $PSScriptRoot
$Hoofd = Split-Path $Bron
$Basis = Join-Path $env:TEMP 'opioidrotatie_basis.xlsx'
$Bas = Join-Path $env:TEMP 'Pharmacom_import.bas'
$Doel = Join-Path $Hoofd 'Opioidrotatie_morfinepomp.xlsm'

python -I (Join-Path $Bron 'werkmap_opbouw.py') $Basis
if ($LASTEXITCODE -ne 0) { throw "werkmap_opbouw.py mislukt" }

# VBA-import verwacht ANSI met CRLF
$Code = [IO.File]::ReadAllText((Join-Path $Bron 'Pharmacom.bas'), [Text.Encoding]::UTF8) -replace "`r?`n", "`r`n"
[IO.File]::WriteAllText($Bas, $Code, [Text.Encoding]::GetEncoding(1252))

$X = New-Object -ComObject Excel.Application
$X.DisplayAlerts = $false
try {
    $Wb = $X.Workbooks.Open($Basis)
    [void]$Wb.VBProject.VBComponents.Import($Bas)
    $Ws = $Wb.Worksheets.Item('Rotatie')
    # Knoppen als afgeronde vormen rechts in de blauwe kopbalk (BENU-kleuren; RGB als R + G*256 + B*65536)
    function Kleur($Hex) { [int]('0x' + $Hex.Substring(0, 2)) + 256 * [int]('0x' + $Hex.Substring(2, 2)) + 65536 * [int]('0x' + $Hex.Substring(4, 2)) }
    function Knop($Naam, $Tekst, $Actie, $Links, $Breedte, $Vulling, $TekstKleur, $Rand) {
        $K = $Ws.Shapes.AddShape(5, $Links, $Top, $Breedte, 30)    # 5 = afgeronde rechthoek
        $K.Name = $Naam
        $K.OnAction = $Actie
        $K.Adjustments.Item(1) = 0.25
        $K.Fill.ForeColor.RGB = (Kleur $Vulling)
        if ($Rand) { $K.Line.ForeColor.RGB = (Kleur $Rand); $K.Line.Weight = 1.25 } else { $K.Line.Visible = 0 }
        $K.Shadow.Visible = 0
        $T = $K.TextFrame2
        $T.VerticalAnchor = 3
        $T.MarginLeft = 6; $T.MarginRight = 6
        $T.TextRange.Text = $Tekst
        $T.TextRange.ParagraphFormat.Alignment = 2
        $T.TextRange.Font.Name = 'Segoe UI Semibold'
        $T.TextRange.Font.Size = 10
        $T.TextRange.Font.Fill.ForeColor.RGB = (Kleur $TekstKleur)
        $K.Placement = 3
        $K.Locked = $true
    }
    $Rechts = $Ws.Range('I1').Left + $Ws.Range('I1').Width - 6
    $Top = $Ws.Range('A1').Top + 18
    Knop 'KnopWis' 'Wis invoer' 'WisInvoer' ($Rechts - 92) 92 'FFFFFF' '002664' $null
    Knop 'KnopLees' ('Lees uit Pharmacom   ' + [char]0x00B7 + '   Ctrl+Alt+F3') 'LeesUitPharmacom' ($Rechts - 92 - 10 - 230) 230 '58A618' 'FFFFFF' $null
    $Ws.Activate()
    $X.ActiveWindow.ScrollRow = 1
    $Ws.Range('B10').Select()
    $Wb.SaveAs($Doel, 52)
    $Wb.Close($false)
} finally {
    $X.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($X)
}
Remove-Item -LiteralPath $Basis, $Bas -Force
"Gebouwd: $Doel"

# Macro's werken niet vanaf een netwerkschijf in de internetzone; gebruik daarom de
# lokale kopie die Installeer.ps1 neerzet.
if ($Installeer) { & (Join-Path $Hoofd 'Installeer.ps1') -Stil }
