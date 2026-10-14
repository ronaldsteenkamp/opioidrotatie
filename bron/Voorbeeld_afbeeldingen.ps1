# Maakt PNG-voorbeelden van alle tabbladen (via PDF-export van Excel + PyMuPDF), om de opmaak te controleren.
# Gebruik: powershell -File Voorbeeld_afbeeldingen.ps1 -Uit <map>
param([Parameter(Mandatory = $true)][string]$Uit)
$ErrorActionPreference = 'Stop'
$Pad = Join-Path (Split-Path $PSScriptRoot) 'Opioidrotatie_morfinepomp.xlsm'
New-Item -ItemType Directory -Force $Uit | Out-Null
$X = New-Object -ComObject Excel.Application
$X.DisplayAlerts = $false
try {
    $Wb = $X.Workbooks.Open($Pad, 0, $true)
    foreach ($Ws in $Wb.Worksheets) {
        $Pdf = Join-Path $Uit ($Ws.Name + '.pdf')
        $Ws.ExportAsFixedFormat(0, $Pdf)
    }
    $Wb.Close($false)
} finally {
    $X.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($X)
}
python -I -c "import pymupdf,glob,os,sys
for f in glob.glob(os.path.join(sys.argv[1],'*.pdf')):
    d=pymupdf.open(f)
    for i,p in enumerate(d):
        p.get_pixmap(dpi=110).save(f[:-4]+'_%d.png'%(i+1))
    d.close(); os.remove(f)
print('ok')" $Uit
