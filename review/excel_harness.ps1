# Haalt de casussen uit casussen.json door de echte werkmap (Excel via COM, alleen-lezen geopend,
# nooit opgeslagen) en schrijft de uitkomsten naar werkmap_uitkomsten.json.
# Ook: parsercontrole via de VBA-testfuncties VertaalCodeTekst/SterkteTekst en Verwerk (verzonnen regels).
# Gebruik (Windows, Excel):  powershell -ExecutionPolicy Bypass -File review\excel_harness.ps1
$ErrorActionPreference = 'Stop'
$Hier = $PSScriptRoot
$Pad = Join-Path (Split-Path $Hier) 'Opioidrotatie_morfinepomp.xlsm'
$Cas = [IO.File]::ReadAllText((Join-Path $Hier 'casussen.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$M = "'Opioidrotatie_morfinepomp.xlsm'!"

function Zet1($Cel, $v) { if ($v -is [string]) { $Cel.Value2 = [string]$v } else { $Cel.Value2 = [double]$v } }
function W($Cel) { $v = $Cel.Value2; if ($v -is [int] -and $v -lt -2146820000) { return "#ERR:" + $Cel.Text }; return $v }

$X = New-Object -ComObject Excel.Application
$X.DisplayAlerts = $false
$Uit = [ordered]@{ pomp = @(); omzet = @(); codes = @(); sterktes = @(); verwerk = @(); extra = [ordered]@{} }
try {
    $Wb = $X.Workbooks.Open($Pad, 0, $true)     # alleen-lezen
    $Ws = $Wb.Worksheets.Item('Rotatie'); $Zs = $Wb.Worksheets.Item('Omzetten')

    function Zet($C) {
        $Ws.Range('B10:D15').ClearContents(); $Ws.Range('D20:D22').ClearContents()
        $r = 10
        foreach ($Rg in $C.regels) { $Ws.Range("B$r").Value2 = 'Onderhoud'; $Ws.Range("C$r").Value2 = [string]$Rg[0]; Zet1 $Ws.Range("D$r") $Rg[1]; $r++ }
        $Ws.Range('D19').Value2 = [string]$C.reden
        if ($null -ne $C.extra) { $Ws.Range('D21').Value2 = [double]$C.extra }
        if ($null -ne $C.handmatig) { $Ws.Range('D20').Value2 = [double]$C.handmatig }
        if ($null -ne $C.egfr) { Zet1 $Ws.Range('D22') $C.egfr }
        if ($null -ne $C.conc) { $Ws.Range('D26').Value2 = [double]$C.conc; $Ws.Range('D27').Value2 = [double]$C.stap; $Ws.Range('D28').Value2 = [double]$C.volume }
        $X.Calculate()
    }

    foreach ($C in $Cas.pomp) {
        Zet $C
        $R = [ordered]@{ id = $C.id }
        foreach ($i in 41..56) { $R["D$i"] = W $Ws.Range("D$i") }
        $R.G16 = W $Ws.Range('G16')
        $R.tegels = @('B33', 'D33', 'G33', 'I33' | ForEach-Object { $Ws.Range($_).Text })
        $R.controle = @(10..15 | ForEach-Object { $Ws.Range("H$_").Text } | Where-Object { $_ })
        $R.waarschuwingen = @(60..72 | ForEach-Object { $Ws.Range("C$_").Text } | Where-Object { $_ })
        $Uit.pomp += $R
    }
    $Ws.Range('D26').Value2 = 10; $Ws.Range('D27').Value2 = 0.01; $Ws.Range('D28').Value2 = 50

    foreach ($C in $Cas.omzet) {
        $C | Add-Member -NotePropertyName conc -NotePropertyValue $null -Force
        Zet $C
        $Zs.Range('C19').Value2 = [string]$C.doel
        if ($null -ne $C.freq) { $Zs.Range('C20').Value2 = [double]$C.freq } else { $Zs.Range('C20').ClearContents() }
        $X.Calculate()
        $R = [ordered]@{ id = $C.id }
        foreach ($a in 'C14', 'C15', 'C16', 'C21', 'C22', 'K25', 'L25', 'P26', 'Q25', 'Q26', 'C42') { $R[$a] = W $Zs.Range($a) }
        foreach ($a in 'C25', 'C26', 'C43', 'B51') { $R[$a] = $Zs.Range($a).Text }
        $R.tabel = @(28..39 | ForEach-Object { $q = $_; if ($Zs.Range("B$q").Text) { (@('B', 'C', 'D', 'E', 'F', 'G', 'H') | ForEach-Object { $Zs.Range("$_$q").Text }) -join ' | ' } } | Where-Object { $_ })
        $Uit.omzet += $R
    }
    $Zs.Range('C20').ClearContents()

    # Preparaten leeg: wat zegt Omzetten?
    $Ps = $Wb.Worksheets.Item('Preparaten')
    Zet $Cas.omzet[2]; $Zs.Range('C19').Value2 = 'Oxycodon oraal'
    $Ps.Range('B7:J86').ClearContents(); $X.Calculate()
    $Uit.extra.preparaten_leeg = [ordered]@{ C25 = $Zs.Range('C25').Text; C26 = $Zs.Range('C26').Text; C43 = $Zs.Range('C43').Text; B51 = $Zs.Range('B51').Text }

    # Eigen doseercodes: zoals het tabblad ze laat invullen (kolom B..F vanaf rij 9) en zoals de VBA ze leest (A7:E56)
    $Ds = $Wb.Worksheets.Item('Doseercodes')
    $Ds.Range('B10').Value2 = 'EIGEN1'; $Ds.Range('C10').Value2 = 3; $Ds.Range('D10').Value2 = 'N'
    $Uit.extra.eigen_code_volgens_tabblad = $X.Run($M + 'VertaalCodeTekst', 'EIGEN1')
    $Uit.extra.voorbeeld_VOORB1 = $X.Run($M + 'VertaalCodeTekst', 'VOORB1')
    $Ds.Range('B10:D10').ClearContents()
    $Ds.Range('A12').Value2 = 'EIGEN2'; $Ds.Range('B12').Value2 = 3; $Ds.Range('C12').Value2 = 'N'
    $Uit.extra.eigen_code_in_kolom_A = $X.Run($M + 'VertaalCodeTekst', 'EIGEN2')
    $Ds.Range('A12:C12').ClearContents()

    foreach ($K in $Cas.codes) { $Uit.codes += [ordered]@{ code = $K; vba = $X.Run($M + 'VertaalCodeTekst', [string]$K) } }
    foreach ($N in $Cas.sterktes) { $Uit.sterktes += [ordered]@{ naam = $N[0]; atc = $N[1]; vba = $X.Run($M + 'SterkteTekst', [string]$N[0], [string]$N[1]) } }

    # Verwerk: verzonnen regels end-to-end (Stil = geen meldingen)
    foreach ($V in $Cas.verwerk) {
        $Kop = "OK`tPharmacom - Dossier - FICTIEVE TESTPATIENT"
        if ($V.egfr) { $Kop += "`t" + $V.egfr[0] + "`t" + $V.egfr[1] + "`tCKD-EPI, eGFR" }
        $T = $Kop + "`nSrt`tEtiketnaam`tAfl. Hoev.`tDosering`tTh. einddatum`tActief`tStartdatum`tATC`n"
        foreach ($L in $V.regels) { $T += "G`t" + $L[0] + "`t1 ST`t" + $L[1] + "`t" + $L[3] + "`t`t01-10-2026`t" + $L[2] + "`n" }
        $Melding = $X.Run($M + 'Verwerk', $T, $true)
        $X.Calculate()
        $R = [ordered]@{ id = $V.id; melding = $Melding; D22 = W $Ws.Range('D22'); D22_oranje = ($Ws.Range('D22').Interior.Color -eq 15069695); regels = @() }
        foreach ($q in 10..15) {
            if ($Ws.Range("C$q").Text -or $Ws.Range("I$q").Text) {
                $R.regels += [ordered]@{ soort = $Ws.Range("B$q").Text; middel = $Ws.Range("C$q").Text; dosis = (W $Ws.Range("D$q")); ome = (W $Ws.Range("G$q"));
                    oranje = ($Ws.Range("I$q").Interior.Color -eq 15069695); uitleg = $Ws.Range("I$q").Text }
            }
        }
        $R.waarschuwingen = @(60..72 | ForEach-Object { $Ws.Range("C$_").Text } | Where-Object { $_ })
        $Uit.verwerk += $R
    }
    $Wb.Close($false)
} finally {
    $X.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($X)
}
$Json = $Uit | ConvertTo-Json -Depth 8
[IO.File]::WriteAllText((Join-Path $Hier 'werkmap_uitkomsten.json'), $Json, (New-Object Text.UTF8Encoding $false))
"geschreven: werkmap_uitkomsten.json"
