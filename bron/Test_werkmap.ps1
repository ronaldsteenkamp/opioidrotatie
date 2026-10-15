# Test de VBA-vertaler en sterkteherkenning in Opioidrotatie_morfinepomp.xlsm (geen patientgegevens).
# De route Pharmacom -> Excel test je met: Pharmacom_naar_Excel.ps1 -Proef -Stil (32-bit PowerShell).
$ErrorActionPreference = 'Stop'
$Pad = Join-Path (Split-Path $PSScriptRoot) 'Opioidrotatie_morfinepomp.xlsm'
$M = "'Opioidrotatie_morfinepomp.xlsm'!"
$X = New-Object -ComObject Excel.Application
$X.DisplayAlerts = $false
try {
    $Wb = $X.Workbooks.Open($Pad)
    "--- doseercodes"
    foreach ($C in 'P1CM6', '2.1T', '3.2T', '23.1T', '1.23T', 'MA1T', 'M2A1T', 'MP1S', 'NP1C', 'NH1T', 'D31PL', '1W1PL',
        'D41PL', 'D42PL', 'Z1M4', 'Z12M8', '13P1T', '13Z5M', '12.%M', '4.5ML', 'OD1T', '3U1T', 'ZO-1T', 'VOORB1', 'VS', '1HOE4', 'FENTN') {
        "{0,-7} -> {1}" -f $C, $X.Run($M + 'VertaalCodeTekst', $C)
    }
    "--- sterkte"
    $Namen = @(
        @('OXYCODON CAPS 10MG', 'N02AA05'), @('OXYCODON HCL TABLET MGA 40MG', 'N02AA05'),
        @('MORFINE HCL DRANK 2MG/ML', 'N02AA01'), @('MORFINE INJVLST 10MG/ML', 'N02AA01'),
        @('FENTANYL PLEISTER 25UG/UUR', 'N02AB03'), @('FENTANYL NEUSSPRAY 100UG/DO', 'N02AB03'),
        @('BUPRENORFINE PLEISTER 10UG/UUR 7D', 'N02AE01'), @('OXYCODON/NALOXON TAB MGA 10/5MG', 'N02AA55'),
        @('PARACETAMOL/CODEINE TABLET 500/10MG', 'N02AJ06'), @('TRAMADOL HCL CAPSULE 50MG', 'N02AX02'),
        @('HYDROMORFON CAPS 2,6MG', 'N02AA03'), @('OXYCODON DRANK 5MG/5ML', 'N02AA05'),
        @('PARACETAMOL TABLET 500MG', 'N02BE01'), @('MORFINE TABLET MGA 30MG', ''))
    foreach ($N in $Namen) { "{0,-38} -> {1}" -f $N[0], $X.Run($M + 'SterkteTekst', $N[0], $N[1]) }
    "--- verwerken (verzonnen regels, geen patientgegevens)"
    $Nep = "OK`tPharmacom - Dossier - FICTIEVE TESTPATIENT`n" +
        "Srt`tEtiketnaam`tAfl. Hoev.`tDosering`tTh. einddatum`tActief`tStartdatum`tATC`n" +
        "G`tOXYCODON TABLET MGA 40MG`t60 ST`t2.1T`t`t`t01-10-2026`tN02AA05`n" +
        "G`tOXYCODON CAPS 10MG`t30 ST`tP1CM6`t`t`t01-10-2026`tN02AA05`n" +
        "G`tPARACETAMOL TABLET 500MG`t100 ST`t3.2T`t`t`t01-10-2026`tN02BE01`n" +
        "G`tFENTANYL PLEISTER 25UG/UUR`t10 ST`tD31PL`t`t`t01-10-2026`tN02AB03`n" +
        "G`tFENTANYL NEUSSPRAY 100UG/DO`t1 ST`tFENTN`t`t`t01-10-2026`tN02AB03`n" +
        "G`tMORFINE HCL DRANK 2MG/ML`t100 ML`t4.5ML`t`t`t01-10-2026`tN02AA01`n" +
        "G`tTRAMADOL HCL CAPSULE 50MG`t30 ST`t23.1C`t`t`t01-10-2026`tN02AX02`n" +
        "G`tOXYCODON CAPS 5MG`t20 ST`t3.1C`t01-01-2020`t`t01-12-2019`tN02AA05`n" +
        "G`tHYDROMORFON CAPS 2,6MG`t20 ST`tVS`t`t`t01-10-2026`tN02AA03`n"
    $X.Run($M + 'Verwerk', $Nep, $true)
    $Ws = $Wb.Worksheets.Item('Rotatie')
    "eGFR na uitlezen zonder nierfunctie (moet leeg): [" + $Ws.Range('D22').Text + "]"
    $X.CalculateFull()
    "B4: " + $Ws.Range('B4').Text
    foreach ($R in 10..15) {
        if ($Ws.Range("C$R").Text) { "rij {0}: {1} | {2} | {3} {4} | OME {5} | oranje={6}`n        {7}" -f $R, $Ws.Range("B$R").Text, $Ws.Range("C$R").Text, $Ws.Range("D$R").Text, $Ws.Range("E$R").Text, $Ws.Range("G$R").Text, ($Ws.Range("I$R").Interior.Color -eq 15069695), $Ws.Range("I$R").Text }
    }
    "Totaal OME: " + $Ws.Range('G16').Text + "  |  tegels: " + $Ws.Range('B33').Text + " / " + $Ws.Range('D33').Text + " / " + $Ws.Range('G33').Text + " / " + $Ws.Range('I33').Text
    "--- verwerken: methadon (staffel) en snelwerkend fentanyl"
    $Nep2 = "OK`tPharmacom - Dossier - FICTIEVE TESTPATIENT 2`t>90`t01-03-2026`tCKD-EPI, eGFR`n" +
        "Srt`tEtiketnaam`tAfl. Hoev.`tDosering`tTh. einddatum`tActief`tStartdatum`tATC`n" +
        "G`tMETHADON HCL TABLET 5MG`t30 ST`t2.2T`t`t`t01-10-2026`tN07BC02`n" +
        "G`tABSTRAL SUBLINGUALE TABLET 100MICROGRAM`t10 ST`tP1TM4`t`t`t01-10-2026`tN02AB03`n" +
        "G`tFENTANYL NEUSSPRAY 100UG/DO`t1 ST`tFENTN`t`t`t01-10-2026`tN02AB03`n"
    $X.Run($M + 'Verwerk', $Nep2, $true)
    $X.CalculateFull()
    "eGFR (D22): " + $Ws.Range('D22').Text + " | oranje=" + ($Ws.Range('D22').Interior.Color -eq 15069695) + " | " + $Ws.Range('D22').Comment.Text()
    foreach ($R in 10..15) {
        if ($Ws.Range("C$R").Text) { "rij {0}: {1} | {2} | {3} {4} | factor {5} | OME {6} | {7} | oranje={8}" -f $R, $Ws.Range("B$R").Text, $Ws.Range("C$R").Text, $Ws.Range("D$R").Text, $Ws.Range("E$R").Text, $Ws.Range("F$R").Text, $Ws.Range("G$R").Text, $Ws.Range("H$R").Text, ($Ws.Range("I$R").Interior.Color -eq 15069695) }
    }
    "methadonfactor bij 20 / 21 / 50 / 51 mg:"
    foreach ($Dos in 20, 21, 50, 51) { $Ws.Range('C15').Value2 = 'Methadon oraal'; $Ws.Range('D15').Value2 = $Dos; $X.Calculate(); "  {0} mg -> factor {1}, OME {2}" -f $Dos, $Ws.Range('F15').Text, $Ws.Range('G15').Text }
    $Qs = $Wb.Worksheets.Item('Omrekentabel')
    $Kol = 3; while ($Qs.Cells.Item(5, $Kol).Text -and $Qs.Cells.Item(5, $Kol).Text -ne 'Methadon oraal') { $Kol++ }
    "Omrekentabel methadon (OME -> mg): " + ((7..16 | ForEach-Object { $Qs.Cells.Item($_, 2).Text + '->' + $Qs.Cells.Item($_, $Kol).Text }) -join ', ')
    "Waarschuwingen: " + (($Ws.UsedRange.Cells | Where-Object { $_.Text -like 'Methadon:*' -or $_.Text -like 'Snelwerkend*' } | ForEach-Object { $_.Text.Substring(0, 40) }) -join ' | ')
    "--- omzetten (oxycodon oraal 80 + 30 mg = 165 mg OME)"
    $Ws.Range('B10:D15').ClearContents()
    $Ws.Range('B10').Value2 = 'Onderhoud'; $Ws.Range('C10').Value2 = 'Oxycodon oraal'; $Ws.Range('D10').Value2 = '80'
    $Ws.Range('B11').Value2 = 'Doorbraak'; $Ws.Range('C11').Value2 = 'Oxycodon oraal'; $Ws.Range('D11').Value2 = '30'
    $Ws.Range('D19').Value2 = 'Onvoldoende pijnstilling'
    $Zs = $Wb.Worksheets.Item('Omzetten')
    foreach ($Doel in 'Fentanyl pleister', 'Morfine oraal', 'Hydromorfon oraal', 'Buprenorfine pleister', 'Tapentadol oraal', 'Methadon oraal', 'Fentanyl nasaal', 'Piritramide s.c./i.v.') {
        $Zs.Range('C19').Value2 = [string]$Doel; $X.Calculate()
        "{0}: {1} {2}`n  1 sterkte : {3}`n  combinatie: {4}`n  doorbraak : {5}`n  groen     : {6}" -f $Doel, $Zs.Range('C22').Text, $Zs.Range('D22').Text, $Zs.Range('C25').Text, $Zs.Range('C26').Text, $Zs.Range('C43').Text, $(if ($Zs.Range('C26').DisplayFormat.Font.Bold) { 'combinatie' } elseif ($Zs.Range('C25').DisplayFormat.Font.Bold) { '1 sterkte' } else { '-' })
        $W = 51..60 | ForEach-Object { $Zs.Range("B$_").Text } | Where-Object { $_ -and $_ -notlike 'Rekenhulp*' }
        if ($W) { "  waarschuwing: " + ($W -join ' | ') }
    }
    $Zs.Range('C19').Value2 = 'Oxycodon oraal'; $Zs.Range('C20').Value2 = '3'; $Ws.Range('C10').Value2 = 'Morfine oraal'; $Ws.Range('D10').Value2 = '45'; $Ws.Range('B11:D11').ClearContents(); $X.Calculate()
    "Morfine 45 mg -> oxycodon 3x per dag: " + $Zs.Range('C25').Text + "  |  " + $Zs.Range('C26').Text
    $Wb.Close($false)
} finally {
    $X.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($X)
}
