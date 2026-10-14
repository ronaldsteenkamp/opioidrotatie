"""Bouwt de werkmap Opioïdrotatie (basis-xlsx; VBA en knoppen komen erbij via Bouw_werkmap.ps1).

Opmaak in BENU-huisstijl: koningsblauw #002664, groen #58A618 / #407C05, lichtgrijsblauw #EEF2F7.
Lettertype Segoe UI (Ubuntu, het BENU-webfont, staat niet standaard op Windows).

LET OP: Pharmacom.bas gebruikt de cellen B4 (patiëntregel) en B/C/D/I van de rijen
R0..R1 (blok 1). Verplaats je die, pas dan R_EERST/R_LAATST in Pharmacom.bas aan.
"""
import sys
from copy import copy

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "Opioidrotatie_basis.xlsx"
VERSIE = "Versie 1.4 - 14-10-2026"

# ------------------------------------------------------------------ huisstijl
BLAUW = "002664"        # BENU royal blue
GROEN = "58A618"        # BENU kelly green
GROEN_D = "407C05"      # BENU accessibility green
GROEN_L = "EEF9E5"      # BENU alabaster
LICHT = "EEF2F7"        # BENU anti-flash
LIJN = "D7DEE9"         # BENU general gray
GRIJS_L = "F5F7F9"      # BENU grey lightest
GRIJS = "546176"        # BENU black coral
BLAUW_M = "5C77A2"      # BENU blue yonder
BLAUW_P = "AFBED4"      # BENU pastel blue
ORANJE = "C25700"       # BENU accent orange (toegankelijk)
ORANJE_L = "FFF1E5"
WIT = "FFFFFF"

F = "Segoe UI"
FS = "Segoe UI Semibold"


def font(size=10, color=GRIJS, bold=False, italic=False, name=F):
    return Font(name=name, size=size, color=color, bold=bold, italic=italic)


def fill(kleur):
    return PatternFill("solid", fgColor=kleur)


def cf_fill(kleur):
    """Vulling voor voorwaardelijke opmaak (Excel gebruikt daar de achtergrondkleur)."""
    return PatternFill(fill_type="solid", start_color=kleur, end_color=kleur)


def side(kleur=LIJN, stijl="thin"):
    return Side(style=stijl, color=kleur)


f_norm = font()
f_label = font(color=BLAUW)
f_klein = font(9, GRIJS, italic=True)
f_invoer = font(10, BLAUW, bold=True)
f_kop_tabel = font(10, WIT, name=FS)
f_waarde = font(10, BLAUW)
f_sterk = font(10, BLAUW, name=FS)

rand_invoer = Border(left=side(BLAUW_M), right=side(BLAUW_M), top=side(BLAUW_M), bottom=side(BLAUW_M))
rand_cel = Border(bottom=side(LIJN))
links = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
midden = Alignment(horizontal="center", vertical="center", wrap_text=True)
rechts = Alignment(horizontal="right", vertical="center", indent=1)
boven = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)

wb = Workbook()
try:  # standaardstijl van de werkmap: Segoe UI 10
    wb._named_styles["Normal"].font = Font(name=F, size=10, color=GRIJS)
except Exception:
    pass


def kopbalk(sh, titel, sub, laatste_kol, hoogte_titel=40):
    """Blauwe kopbalk (rij 1-2) met groene accentlijn (rij 3) over kolommen A..laatste_kol."""
    kols = [chr(c) for c in range(ord("A"), ord(laatste_kol) + 1)]
    for c in kols:
        sh[f"{c}1"].fill = fill(BLAUW)
        sh[f"{c}2"].fill = fill(BLAUW)
        sh[f"{c}3"].fill = fill(GROEN)
    sh["B1"] = titel
    sh["B1"].font = Font(name=FS, size=20, color=WIT)
    sh["B1"].alignment = Alignment(vertical="bottom")
    sh["B2"] = sub
    sh["B2"].font = Font(name=F, size=10, color=BLAUW_P)
    sh["B2"].alignment = Alignment(vertical="top")
    sh.row_dimensions[1].height = hoogte_titel
    sh.row_dimensions[2].height = 22
    sh.row_dimensions[3].height = 4
    sh.sheet_view.showGridLines = False


def tabelkop(sh, rij, kol0, koppen, hoogte=30):
    for j, h in enumerate(koppen):
        c = sh.cell(row=rij, column=kol0 + j, value=h)
        c.font = f_kop_tabel
        c.fill = fill(BLAUW)
        c.alignment = midden
    sh.row_dimensions[rij].height = hoogte


def invoer(c, waarde=None, fmt=None):
    if waarde is not None:
        c.value = waarde
    c.font = f_invoer
    c.fill = fill(WIT)
    c.border = rand_invoer
    c.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        c.number_format = fmt


# ================================================================== Omrekenfactoren
fs = wb.active
fs.title = "Omrekenfactoren"
ws = wb.create_sheet("Rotatie", 0)
for col, w in zip("ABCDEFGH", [2, 34, 13, 34, 13, 14, 44, 56]):
    fs.column_dimensions[col].width = w
kopbalk(fs, "Omrekenfactoren", "Volgens KNMP Informatorium Medicamentorum (Analgetica, G806). Witte velden zijn aanpasbaar; "
        "alle berekeningen gebruiken deze waarden.", "I")

fs["B5"] = "Instellingen"
fs["B5"].font = Font(name=FS, size=13, color=BLAUW)
fs.row_dimensions[5].height = 24
settings = [
    ("Verhouding morfine oraal : morfine s.c./i.v. (x : 1)", 3, "0.0",
     "KNMP IM Analgetica - opioïdrotatietabel (uit NVA-richtlijn Pijn bij kanker): 30 mg morfine oraal = 10 mg morfine s.c./i.v."),
    ("Dosis bij rotatie wegens bijwerkingen", 0.75, "0%",
     "KNMP IM Analgetica, Opioïdrotatie: bij rotatie vanwege bijwerkingen starten met 75% van de equivalente 24-uursdosering."),
    ("Dosis bij rotatie wegens onvoldoende pijnstilling", 1.0, "0%",
     "KNMP IM Analgetica: bij onvoldoende effect het nieuwe opioïd starten met een equivalente dosering."),
    ("Drempel 'hoge dosis' (OME mg/24 u) voor extra waarschuwing", 240, "0",
     "Afgeleid van KNMP: bij de meeste patiënten met kanker is < 240 mg morfine oraal/dag voldoende. Boven deze dosis extra alert."),
    ("Doorbraakbolus (fractie van dagdosis)", 1 / 6, "0.0%",
     "KNMP IM Analgetica, Doorbraakpijn: 1/6 van de (equivalente) dagdosis per keer, max. 4-6x per dag; "
     ">3x per dag nodig → onderhoudsdosering met 50% verhogen."),
    ("Doorbraakbolus - lagere variant (fractie van dagdosis)", 0.10, "0%",
     "Pallialine-pijnkaart: 10-15% van de dagdosis. Ter vergelijking; KNMP hanteert 1/6."),
    ("Beschikbare concentratie morfine 1 (mg/ml)", 10, "0.0#",
     "Eigen voorraad: morfine 10 mg/ml voor de pomp. Verschijnt in de keuzelijst op tabblad Rotatie."),
    ("Beschikbare concentratie morfine 2 (mg/ml)", 1, "0.0#",
     "Eigen voorraad: morfine 1 mg/ml, voor lage doseringen. Verschijnt in de keuzelijst op tabblad Rotatie."),
]
S0 = 6
for i, (lab, val, fmt, src) in enumerate(settings):
    r = S0 + i
    fs[f"B{r}"] = lab
    fs[f"B{r}"].font = f_label
    fs[f"B{r}"].alignment = links
    invoer(fs[f"C{r}"], val, fmt)
    fs[f"D{r}"] = src
    fs[f"D{r}"].font = f_klein
    fs[f"D{r}"].alignment = links
    fs.merge_cells(f"D{r}:H{r}")
    for c in "BDEFGH":
        fs[f"{c}{r}"].border = rand_cel
    fs.row_dimensions[r].height = 30
S_RATIO, S_SIDE, S_INSUF, S_HIGH, S_BOLLO, S_BOLHI, S_CONC1, S_CONC2 = [f"Omrekenfactoren!$C${S0 + i}" for i in range(8)]
CONC_LIJST = f"Omrekenfactoren!$C${S0 + 6}:$C${S0 + 7}"

fs["B15"] = "Omrekenfactoren naar OME (orale morfine-equivalenten, mg/24 u)"
fs["B15"].font = Font(name=FS, size=13, color=BLAUW)
fs.row_dimensions[15].height = 24
tabelkop(fs, 16, 2, ["Middel en toedieningsweg", "Eenheid invoer", "Equianalgetische verhouding",
                     "Factor → OME", "Max. dagdosis (invoereenheid)", "Bron", "Opmerking"], 34)

KNMP = "KNMP IM Analgetica (G806), Opioïdrotatie - conversiefactoren en rotatietabel (bron: NVA-richtlijn Pijn bij kanker)"
KNMPT = "KNMP IM Analgetica (G806), opioïdrotatietabel"
OTH = "Niet in KNMP-rotatietabel; gangbare literatuurwaarde"
drugs = [
    ("Morfine oraal", "mg/24 u", "Referentie", 1, None, KNMP, ""),
    ("Morfine s.c./i.v.", "mg/24 u", "Morfine oraal : s.c./i.v. = 3 : 1 (zie instelling)", f"={S_RATIO}", None, KNMPT,
     "Volgt de instelling hierboven. KNMP-tabel: 60 mg oraal = 20 mg s.c./i.v."),
    ("Oxycodon oraal", "mg/24 u", "Morfine oraal : oxycodon oraal = 1,5 : 1", 1.5, None, KNMP,
     "KNMP-tabel: 60 mg morfine oraal = 40 mg oxycodon oraal."),
    ("Oxycodon s.c./i.v.", "mg/24 u", "Oxycodon oraal : s.c./i.v. = 2 : 1", 3, None, KNMPT,
     "KNMP-tabel: 60 mg morfine oraal = 20 mg oxycodon s.c./i.v."),
    ("Hydromorfon oraal", "mg/24 u", "Morfine oraal : hydromorfon oraal = 5 : 1", 5, None, KNMP,
     "KNMP-tabel: 60 mg morfine oraal = 12 mg hydromorfon oraal."),
    ("Hydromorfon s.c./i.v.", "mg/24 u", "Hydromorfon oraal : s.c./i.v. = 3 : 1", 15, None, KNMPT,
     "KNMP-tabel: 60 mg morfine oraal = 4 mg hydromorfon s.c./i.v."),
    ("Fentanyl pleister", "µg/uur", "Morfine oraal : fentanyl TD = 100 : 1", 2.4, None, KNMP,
     "µg/uur x 24 / 1000 x 100. KNMP-tabel: 25 µg/uur = 60 mg morfine oraal."),
    ("Fentanyl s.c./i.v. continu", "µg/uur", "Morfine parenteraal : fentanyl parenteraal = 100 : 1", 7.2, None,
     "KNMP IM Analgetica (G806): 'in een aantal bronnen 100:1 (beperkte gegevens)'",
     "µg/uur x 24 / 1000 x 100 x 3. LET OP: in de praktijk wordt ook 1:1 met de pleister (factor 2,4) gehanteerd; "
     "dat geeft een 3x lagere morfinedosis. Altijd overleg."),
    ("Buprenorfine pleister", "µg/uur", "Morfine oraal : buprenorfine TD = 100 : 1", 2.4, 140, KNMP,
     "KNMP-tabel: 52,5 µg/uur = 120 mg morfine oraal. Boven 140 µg/uur niet geadviseerd (NVA)."),
    ("Tramadol oraal", "mg/24 u", "Morfine oraal : tramadol = 1 : 5 (KNMP-tabel)", 0.2, 400, KNMPT,
     "KNMP-tabel: 150 mg tramadol = 30 mg, 300 mg = 60 mg morfine oraal. (NVA COPD/hartfalen noemt 1 : 10.) Max. 400 mg/24 u."),
    ("Tapentadol oraal", "mg/24 u", "Morfine oraal : tapentadol = 1 : 2,5", 0.4, 500, KNMP,
     "KNMP-tabel: 150 mg tapentadol = 60 mg morfine oraal. Max. onderzochte dosis 500 mg/24 u."),
    ("Piritramide s.c./i.v.", "mg/24 u", "15-20 mg piritramide = 10 mg morfine parenteraal", 1.5, None,
     "KNMP IM Analgetica (G806), equivalentie piritramide",
     "Gerekend met 20 : 10 (voorzichtige kant bij rotatie naar morfine). Bij 15 : 10 wordt de factor 2,0."),
    ("Codeïne oraal", "mg/24 u", "Morfine oraal : codeïne = 1 : 10", 0.1, 240, OTH,
     "Niet in KNMP-rotatietabel; NHG raadt codeïne af als pijnstiller. Verifiëren."),
    ("Methadon oraal", "mg/24 u", "Variabel, 5 : 1 tot 10 : 1 (dosisafhankelijk)", None, None, KNMP,
     "GEEN automatische omrekening: altijd overleg met deskundige (consultatieteam palliatieve zorg / pijnteam)."),
]
D0 = 17
for i, row in enumerate(drugs):
    r = D0 + i
    band = fill(GRIJS_L) if i % 2 else fill(WIT)
    for j, v in enumerate(row):
        c = fs.cell(row=r, column=2 + j, value=v)
        c.font = f_norm
        c.fill = band
        c.alignment = links
        c.border = rand_cel
    fs.cell(row=r, column=2).font = f_sterk
    fs.cell(row=r, column=3).alignment = midden
    fs.cell(row=r, column=7).font = f_klein
    fs.cell(row=r, column=8).font = font(9, GRIJS)
    for col in (5, 6):
        c = fs.cell(row=r, column=col)
        is_formule = isinstance(c.value, str) and c.value.startswith("=")
        invoer(c, fmt="0.0##" if col == 5 else "0")
        if is_formule:
            c.font = f_waarde
            c.fill = fill(LICHT)
    fs.row_dimensions[r].height = 32
D1 = D0 + len(drugs) - 1
NAMES = f"Omrekenfactoren!$B${D0}:$B${D1}"
UNITS = f"Omrekenfactoren!$C${D0}:$C${D1}"
FACT = f"Omrekenfactoren!$E${D0}:$E${D1}"
MAXD = f"Omrekenfactoren!$F${D0}:$F${D1}"
fs.freeze_panes = "A4"
fs.page_setup.orientation = "landscape"
fs.page_setup.fitToWidth = 1
fs.page_setup.fitToHeight = 0
fs.sheet_properties.pageSetUpPr.fitToPage = True
fs.sheet_properties.tabColor = GROEN

# ================================================================== Rotatie
for col, w in zip("ABCDEFGHIJ", [2, 15, 28, 23, 11, 12, 14, 30, 56, 2]):
    ws.column_dimensions[col].width = w
kopbalk(ws, "Opioïdrotatie", "naar morfine s.c./i.v. continu (pomp)   ·   omrekening volgens KNMP Informatorium Medicamentorum",
        "J", 44)
ws.sheet_view.showRowColHeaders = False
ws.sheet_view.zoomScale = 100
ws.sheet_properties.tabColor = BLAUW

# Patiëntregel (B4; wordt door de macro gevuld)
PLAATS = "Nog niet uitgelezen uit Pharmacom  ·  open het dossier en druk op Ctrl+Alt+F3"
ws.merge_cells("B4:I4")
ws["B4"] = PLAATS
ws["B4"].font = Font(name=FS, size=10, color=GROEN_D)
ws["B4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
for c in "BCDEFGHI":
    ws[f"{c}4"].fill = fill(GROEN_L)
ws.row_dimensions[4].height = 26

# Legenda
ws["B5"] = "Legenda"
ws["B5"].font = f_klein
ws["B5"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
for ref, tekst, f_, fi, br in (("C5", "Invulveld", f_invoer, WIT, rand_invoer),
                               ("D5", "Berekend", f_waarde, LICHT, None),
                               ("E5", "Let op", font(9, ORANJE, name=FS), ORANJE_L, None)):
    c = ws[ref]
    c.value = tekst
    c.font = Font(name=f_.name, size=9, color=f_.color, bold=f_.bold)
    c.fill = fill(fi)
    c.alignment = midden
    if br:
        c.border = br
ws.merge_cells("E5:F5")
ws.row_dimensions[5].height = 20
ws.row_dimensions[6].height = 8


def stap(r, nummer, titel, sub=None):
    """Kop van een stap: groen nummer + blauwe titel, groene lijn eronder."""
    c = ws[f"B{r}"]
    c.value = nummer
    c.font = Font(name=FS, size=12, color=WIT)
    c.fill = fill(GROEN)
    c.alignment = midden
    ws[f"C{r}"] = titel
    ws[f"C{r}"].font = Font(name=FS, size=13, color=BLAUW)
    ws[f"C{r}"].alignment = Alignment(vertical="center", indent=1)
    for k in "BCDEFGHI":
        ws[f"{k}{r}"].border = Border(bottom=side(GROEN, "medium"))
    ws.row_dimensions[r].height = 26
    if sub:
        ws[f"C{r + 1}"] = sub
        ws[f"C{r + 1}"].font = f_klein
        ws[f"C{r + 1}"].alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(f"C{r + 1}:I{r + 1}")
        ws.row_dimensions[r + 1].height = 20


def label(r, tekst, noot=None):
    ws[f"B{r}"] = tekst
    ws[f"B{r}"].font = f_label
    ws[f"B{r}"].alignment = links
    ws.merge_cells(f"B{r}:C{r}")
    if noot:
        ws[f"E{r}"] = noot
        ws[f"E{r}"].font = f_klein
        ws[f"E{r}"].alignment = links
        ws.merge_cells(f"E{r}:I{r}")
    ws.row_dimensions[r].height = 22


def berekend(ref, formule, fmt, sterk=False):
    c = ws[ref]
    c.value = formule
    c.number_format = fmt
    c.font = f_sterk if sterk else f_waarde
    c.fill = fill(LICHT)
    c.alignment = rechts


# ---------------------------------------------------------------- 1. huidig gebruik
stap(7, 1, "Huidig opioïdgebruik",
     "Alle opioïden van de afgelopen 24 uur, inclusief gebruikte doorbraakmedicatie. Pleisters en continue fentanyl in µg/uur.")
tabelkop(ws, 9, 2, ["Soort", "Middel + toedieningsweg", "Dosis per 24 u", "Eenheid", "Factor → OME", "OME (mg/24 u)",
                    "Controle", "Uitgelezen uit Pharmacom"], 30)
R0, R1 = 10, 15
dv_soort = DataValidation(type="list", formula1='"Onderhoud,Doorbraak"', allow_blank=True)
dv_mid = DataValidation(type="list", formula1=NAMES, allow_blank=True,
                        error="Kies een middel uit de lijst (tabblad Omrekenfactoren).", showErrorMessage=True)
dv_num = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                        error="Voer een getal ≥ 0 in.", showErrorMessage=True)
for v in (dv_soort, dv_mid, dv_num):
    ws.add_data_validation(v)
for r in range(R0, R1 + 1):
    invoer(ws[f"B{r}"])
    invoer(ws[f"C{r}"])
    ws[f"C{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    invoer(ws[f"D{r}"], fmt="0.0##")
    dv_soort.add(f"B{r}")
    dv_mid.add(f"C{r}")
    dv_num.add(f"D{r}")
    berekend(f"E{r}", f'=IF(C{r}="","",INDEX({UNITS},MATCH(C{r},{NAMES},0)))', "@")
    ws[f"E{r}"].alignment = midden
    berekend(f"F{r}", f'=IF(C{r}="","",IF(INDEX({FACT},MATCH(C{r},{NAMES},0))="","n.v.t.",'
                      f'INDEX({FACT},MATCH(C{r},{NAMES},0))))', "0.0##")
    berekend(f"G{r}", f'=IF(OR(C{r}="",D{r}=""),"",IF(ISNUMBER(F{r}),D{r}*F{r},"overleg"))', "0.0", sterk=True)
    c = ws[f"H{r}"]
    c.value = (f'=IF(OR(C{r}="",D{r}=""),"",IF(NOT(ISNUMBER(F{r})),"Geen omrekening - overleg deskundige",'
               f'IF(AND(ISNUMBER(INDEX({MAXD},MATCH(C{r},{NAMES},0))),'
               f'D{r}>INDEX({MAXD},MATCH(C{r},{NAMES},0))),'
               f'"Boven max. dagdosis ("&INDEX({MAXD},MATCH(C{r},{NAMES},0))&")","✓ OK")))')
    c.font = font(9, GROEN_D, name=FS)
    c.fill = fill(LICHT)
    c.alignment = links
    c = ws[f"I{r}"]
    c.font = font(8.5, GRIJS)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for k in "BCDEFGHI":
        ws[f"{k}{r}"].border = Border(left=ws[f"{k}{r}"].border.left, right=ws[f"{k}{r}"].border.right,
                                      top=ws[f"{k}{r}"].border.top, bottom=side(LIJN) if k in "EFGHI" else side(BLAUW_M))
    ws.row_dimensions[r].height = 34
# Controle in oranje als er iets mis is
ws.conditional_formatting.add(f"H{R0}:H{R1}", FormulaRule(formula=[f'AND(H{R0}<>"",H{R0}<>"✓ OK")'],
                                                          font=Font(color=ORANJE, bold=True), fill=cf_fill(ORANJE_L)))
# voorbeeld
ws[f"B{R0}"], ws[f"C{R0}"], ws[f"D{R0}"] = "Onderhoud", "Oxycodon oraal", 80
ws[f"B{R0 + 1}"], ws[f"C{R0 + 1}"], ws[f"D{R0 + 1}"] = "Doorbraak", "Oxycodon oraal", 30
ws[f"C{R0}"].comment = Comment("Voorbeeld: oxycodon retard 2dd 40 mg = 80 mg/24 u. Overschrijven of 'Wis invoer'.", "Rekenhulp")
ws[f"C{R0 + 1}"].comment = Comment("Voorbeeld: 3x oxycodon IR 10 mg in afgelopen 24 u = 30 mg.", "Rekenhulp")

TR = R1 + 1
ws[f"B{TR}"] = "Totaal orale morfine-equivalenten (OME)"
ws[f"B{TR}"].font = f_sterk
ws[f"B{TR}"].alignment = links
ws.merge_cells(f"B{TR}:F{TR}")
berekend(f"G{TR}", f"=SUM(G{R0}:G{R1})", "0.0", sterk=True)
ws[f"G{TR}"].font = Font(name=FS, size=11, color=BLAUW)
ws[f"H{TR}"] = "mg morfine oraal / 24 u"
ws[f"H{TR}"].font = f_klein
ws[f"H{TR}"].alignment = links
for k in "BCDEFGHI":
    ws[f"{k}{TR}"].border = Border(top=side(BLAUW, "medium"))
ws.row_dimensions[TR].height = 26
ws.row_dimensions[TR + 1].height = 12

# ---------------------------------------------------------------- 2. reden en patiënt
stap(18, 2, "Reden van rotatie en patiënt")
reasons = ["Bijwerkingen", "Onvoldoende pijnstilling", "Anders (handmatig %)"]
dv_reason = DataValidation(type="list", formula1='"' + ",".join(reasons) + '"', allow_blank=False)
ws.add_data_validation(dv_reason)
label(19, "Reden van rotatie", "Onvoldoende pijnstilling → 100%; bijwerkingen → 75% (KNMP; instelbaar op tabblad Omrekenfactoren).")
invoer(ws["D19"], "Onvoldoende pijnstilling")
dv_reason.add("D19")
label(20, "Handmatig % van equianalgetische dosis", "Alleen gebruikt bij 'Anders (handmatig %)'.")
invoer(ws["D20"], None, "0%")
label(21, "Extra reductie (optioneel)", "Bijv. 25% bij een kwetsbare of oudere patiënt of bij sufheid. Komt bovenop de reden.")
invoer(ws["D21"], 0, "0%")
label(22, "eGFR (ml/min/1,73 m²)", "Optioneel. Geeft een waarschuwing bij nierfunctiestoornis.")
invoer(ws["D22"], None, "0")
pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1.5", allow_blank=True,
                     error="Voer een percentage in (bijv. 75%).", showErrorMessage=True)
ws.add_data_validation(pct)
pct.add("D20")
pct.add("D21")
label(23, "Toe te passen % van de equianalgetische dosis")
berekend("D23", f'=IF(D19="Bijwerkingen",{S_SIDE},IF(D19="Onvoldoende pijnstilling",{S_INSUF},'
                f'IF(ISNUMBER(D20),D20,0)))*(1-N(D21))', "0%", sterk=True)
ws["B23"].font = f_sterk
ws.row_dimensions[24].height = 12

# ---------------------------------------------------------------- 3. pomp
stap(25, 3, "Pomp")
label(26, "Concentratie morfine in spuit/cassette (mg/ml)",
      "Keuzelijst: 10 of 1 mg/ml (eigen voorraad; instelbaar op tabblad Omrekenfactoren).")
invoer(ws["D26"], 10, "0.0#")
dv_conc = DataValidation(type="list", formula1=CONC_LIJST, allow_blank=False,
                         error="Kies een concentratie uit de lijst (tabblad Omrekenfactoren).", showErrorMessage=True)
ws.add_data_validation(dv_conc)
dv_conc.add("D26")
label(27, "Kleinste instelstap van de pomp (ml/uur)", "Bijv. 0,01 of 0,1 - afhankelijk van het type pomp.")
invoer(ws["D27"], 0.01, "0.0##")
label(28, "Volume spuit/cassette (ml)", "Voor de looptijd.")
invoer(ws["D28"], 50, "0")
for ref in ("D27", "D28"):
    dv_num.add(ref)
ws.row_dimensions[29].height = 12

# ---------------------------------------------------------------- 4. resultaat: tegels
stap(30, 4, "Resultaat: morfine s.c./i.v. continu")
ws.row_dimensions[31].height = 8
# Detailregels (onder de tegels), eerst vastleggen voor de verwijzingen
DET = 41
det = [
    ("Totaal OME", f"=G{TR}", "0.0", "mg morfine oraal / 24 u"),
    ("Equianalgetisch: morfine s.c./i.v.", f"=D{DET}/{S_RATIO}", "0.0", "mg / 24 u  (OME ÷ verhouding oraal : s.c.)"),
    ("Toegepast percentage", "=D23", "0%", "reden van rotatie × (1 − extra reductie)"),
    ("Nieuwe dagdosis morfine s.c./i.v.", f"=D{DET + 1}*D{DET + 2}", "0.0", "mg / 24 u"),
    ("Dosis per uur", f"=D{DET + 3}/24", "0.00", "mg / uur"),
    ("Inloopsnelheid (exact)", f"=IF(N(D26)>0,D{DET + 4}/D26,0)", "0.000", "ml / uur"),
    ("Inloopsnelheid afgerond op pompstap", f"=IF(N(D27)>0,ROUND(D{DET + 5}/D27,0)*D27,D{DET + 5})", "0.00",
     "ml / uur  ← instellen op de pomp"),
    ("Werkelijke dagdosis bij afgeronde snelheid", f"=D{DET + 6}*D26*24", "0.0", "mg / 24 u"),
    ("Afwijking door afronding", f"=IF(D{DET + 3}>0,D{DET + 7}/D{DET + 3}-1,0)", "+0.0%;-0.0%;0.0%",
     "t.o.v. de berekende dagdosis"),
    ("Volume per 24 uur", f"=D{DET + 6}*24", "0.0", "ml / 24 u"),
    ("Looptijd spuit/cassette", f"=IF(D{DET + 6}>0,D28/D{DET + 6},0)", "0", "uur"),
    ("Doorbraakbolus s.c. - KNMP (1/6 dagdosis)", f"=D{DET + 7}*{S_BOLLO}", "0.0", "mg per keer; max. 4-6x per dag (KNMP)"),
    ("Doorbraakbolus s.c. - lagere variant (10%)", f"=D{DET + 7}*{S_BOLHI}", "0.0",
     "mg per keer (Pallialine 10-15%), bijv. bij een kwetsbare patiënt"),
    ("Doorbraakbolus in ml (KNMP  |  lagere variant)", None, "@", "ml per bolus; interval volgens lokaal protocol"),
    ("Andere concentratie (ter vergelijking)", f"=IF(N(D26)={S_CONC1},{S_CONC2},{S_CONC1})", '0.0#" mg/ml"',
     "de andere sterkte uit de voorraad"),
    ("Pompstand bij andere concentratie",
     f"=IF(AND(N(D{DET + 14})>0,D{DET + 3}>0),"
     f"IF(N(D27)>0,ROUND(D{DET + 4}/D{DET + 14}/D27,0)*D27,D{DET + 4}/D{DET + 14}),0)", "0.00",
     f'=IF(D{DET + 15}>0,"ml / uur  ·  afwijking "&FIXED((D{DET + 15}*D{DET + 14}*24/D{DET + 3}-1)*100,1,TRUE)'
     f'&"%  ·  looptijd "&FIXED(D28/D{DET + 15},0,TRUE)&" uur","ml / uur")'),
]
(R_OME, R_SC, R_PCT, R_DAG, R_UUR, R_EXACT, R_ML, R_ECHT, R_AFW, R_VOL, R_LOOP, R_BOL, R_BOL2, R_BOLML,
 R_ALTC, R_ALTML) = range(DET, DET + 16)

# Tegels (rij 32-36)
tegels = [
    # (kolommen, label, waarde, formaat, subtekst, hoofdtegel)
    ("B", "C", "POMPSTAND", f'=IF(G{TR}>0,D{R_ML},"")', '0.00" ml/uur"',
     f'=IF(G{TR}>0,"bij "&FIXED(D26,1,TRUE)&" mg/ml  ·  stap "&FIXED(D27,2,TRUE)&" ml/uur","Vul stap 1 in")', True),
    ("D", "F", "MORFINE S.C./I.V.", f'=IF(G{TR}>0,D{R_ECHT},"")', '0.0" mg/24 u"',
     f'=IF(G{TR}>0,"= "&FIXED(D{R_ECHT}/24,2,TRUE)&" mg/uur  ·  "&TEXT(D23,"0%")&" van equianalgetisch","")', False),
    ("G", "H", "DOORBRAAKBOLUS", f'=IF(G{TR}>0,D{R_BOL},"")', '0.0" mg s.c."',
     f'=IF(AND(G{TR}>0,N(D26)>0),"= "&FIXED(D{R_BOL}/D26,2,TRUE)&" ml  ·  1/6 dagdosis, max. 4-6× per dag","")', False),
    ("I", "I", "SPUIT / CASSETTE", f'=IF(G{TR}>0,D{R_LOOP},"")', '0" uur looptijd"',
     f'=IF(G{TR}>0,FIXED(D{R_VOL},1,TRUE)&" ml per 24 u  ·  spuit van "&D28&" ml","")', False),
]
for k0, k1, lab, waarde, fmt, sub, hoofd in tegels:
    achter = BLAUW if hoofd else LICHT
    voor = WIT if hoofd else BLAUW
    subkleur = BLAUW_P if hoofd else GRIJS
    for r in range(32, 37):
        for k in range(ord(k0), ord(k1) + 1):
            ws[f"{chr(k)}{r}"].fill = fill(achter)
    ws.merge_cells(f"{k0}32:{k1}32")
    ws.merge_cells(f"{k0}33:{k1}34")
    ws.merge_cells(f"{k0}35:{k1}35")
    ws[f"{k0}32"] = lab
    ws[f"{k0}32"].font = Font(name=FS, size=9, color=GROEN if hoofd else GROEN_D)
    ws[f"{k0}32"].alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
    ws[f"{k0}33"] = waarde
    ws[f"{k0}33"].number_format = fmt
    ws[f"{k0}33"].font = Font(name=FS, size=22 if hoofd else 18, color=voor)
    ws[f"{k0}33"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws[f"{k0}35"] = sub
    ws[f"{k0}35"].font = Font(name=F, size=8.5, color=subkleur)
    ws[f"{k0}35"].alignment = Alignment(horizontal="left", vertical="top", indent=1, wrap_text=True)
for r, h in ((32, 20), (33, 22), (34, 22), (35, 30), (36, 6)):
    ws.row_dimensions[r].height = h
# Witte scheiding tussen de tegels
for k0, *_ in tegels[1:]:
    for r in range(32, 37):
        ws[f"{k0}{r}"].border = Border(left=Side(style="thick", color=WIT))

# Altijd-blok direct onder de tegels
ws.merge_cells("B37:I38")
ws["B37"] = ("Altijd:  start de pomp op het moment van de volgende geplande dosis van het oude middel  ·  zorg voor doorbraakmedicatie "
             "(KNMP: 1/6 dagdosis, max. 4-6× per dag; > 3× per dag nodig → onderhoud +50%)  ·  evalueer pijn, sufheid en ademhaling  ·  "
             "laat berekening en pompinstelling dubbel controleren.")
ws["B37"].font = font(9, BLAUW)
ws["B37"].alignment = Alignment(wrap_text=True, vertical="center", indent=1)
for rr in (37, 38):
    ws[f"B{rr}"].fill = fill(GROEN_L)
    for k in "CDEFGHI":
        ws[f"{k}{rr}"].fill = fill(GROEN_L)
    ws.row_dimensions[rr].height = 18
ws.row_dimensions[39].height = 12

# Detailtabel
ws["B40"] = "Berekening in detail"
ws["B40"].font = Font(name=FS, size=10, color=BLAUW)
ws["B40"].alignment = Alignment(vertical="center", indent=1)
ws.merge_cells("B40:C40")
for k in "BCDEFGHI":
    ws[f"{k}40"].border = Border(bottom=side(LIJN))
ws.row_dimensions[40].height = 22
for i, (lab, fm, fmt, uitleg) in enumerate(det):
    r = DET + i
    label(r, lab, uitleg)
    ws.row_dimensions[r].height = 20
    if fm:
        berekend(f"D{r}", fm, fmt, sterk=r in (R_DAG, R_ML))
    for k in "BCDEFGHI":
        ws[f"{k}{r}"].border = Border(bottom=side(GRIJS_L))
berekend(f"D{R_BOLML}", f'=IF(N(D26)>0,FIXED(D{R_BOL}/D26,2,TRUE)&"  |  "&FIXED(D{R_BOL2}/D26,2,TRUE),"")', "@")
for r in (R_DAG, R_ML):
    ws[f"B{r}"].font = f_sterk
    ws[f"D{r}"].fill = fill(GROEN_L)
    ws[f"D{r}"].font = Font(name=FS, size=10, color=GROEN_D)
R_DET_END = DET + len(det) - 1
ws.row_dimensions[R_DET_END + 1].height = 12

# ---------------------------------------------------------------- 5. controles
CW = R_DET_END + 2
stap(CW, 5, "Controles en waarschuwingen",
     f"Rekenhulp - vervangt geen klinisch oordeel, lokaal protocol of dubbele controle.  Bron: KNMP Informatorium Medicamentorum, "
     f"Analgetica (G806).  {VERSIE}")
ws[f"C{CW + 1}"].font = font(8.5, BLAUW_M, italic=True)
L = f"C{R0}:C{R1}"
warns = [
    f'=IF(COUNTIF({L},"Methadon*")>0,"Methadon wordt niet meegerekend - omrekening is dosisafhankelijk. Overleg altijd met een deskundige.","")',
    f'=IF(COUNTIF({L},"Fentanyl pleister")>0,"Fentanylpleister: na verwijderen wordt nog 12-24 uur fentanyl afgegeven. Start de pomp in overleg (bijv. ca. 12 uur na verwijderen of met een lagere startdosis) en zorg voor doorbraakmedicatie.","")',
    f'=IF(COUNTIF({L},"Fentanyl s.c./i.v. continu")>0,"Fentanyl s.c./i.v.: de KNMP-factor parenteraal morfine : fentanyl 100 : 1 berust op beperkte gegevens. Bij 1 : 1 met de pleister zou de morfinedosis 3x lager uitkomen. Altijd overleg.","")',
    f'=IF(COUNTIF({L},"Buprenorfine pleister")>0,"Buprenorfinepleister: partiële agonist met lange werkingsduur. Plan de overgang naar morfine in overleg.","")',
    '=IF(AND(ISNUMBER(D22),D22<30),"eGFR < 30: morfine en metabolieten stapelen. Morfine is minder geschikt - overweeg een ander opioïd (bijv. fentanyl of hydromorfon) of een fors lagere dosis in overleg.","")',
    f'=IF(G{TR}>={S_HIGH},"Hoge dosis (OME ≥ "&{S_HIGH}&" mg/24 u): de omrekening is onzekerder. Overweeg extra reductie, stapsgewijze rotatie en overleg met het consultatieteam palliatieve zorg of pijnteam.","")',
    f'=IF(COUNTIF(H{R0}:H{R1},"Boven*")>0,"Een of meer doseringen liggen boven de maximale dagdosis - controleer de invoer.","")',
    f'=IF(COUNTIFS(B{R0}:B{R1},"Doorbraak",I{R0}:I{R1},"*MAXIMUM*")>0,"Zo nodig / bij pijn is uit Pharmacom met het MAXIMUM ingevuld. Verlaag naar het werkelijke gebruik als dat minder was; anders wordt de pompdosis te hoog.","")',
    '=IF(AND(D19="Onvoldoende pijnstilling",N(D21)=0),"Volledige equianalgetische dosis (geen reductie). Evalueer effect en bijwerkingen na enkele uren.","")',
    f'=IF(ABS(D{R_AFW})>0.1,"Afronding op de pompstap wijkt "&FIXED(ABS(D{R_AFW})*100,0,TRUE)&"% af bij "'
    f'&FIXED(D26,1,TRUE)&" mg/ml."&IF(AND(D{R_ALTML}>0,ABS(D{R_ALTML}*D{R_ALTC}*24/D{R_DAG}-1)<=0.1),'
    f'" Kies "&FIXED(D{R_ALTC},0,TRUE)&" mg/ml: pompstand "&FIXED(D{R_ALTML},2,TRUE)&" ml/uur.",'
    f'" Overweeg een andere concentratie of pompstap."),"")',
    f'=IF(AND(G{TR}>0,D{R_LOOP}>0,D{R_LOOP}<24),"De spuit/cassette is binnen 24 uur leeg ("&FIXED(D{R_LOOP},0,TRUE)'
    f'&" uur bij "&FIXED(D26,1,TRUE)&" mg/ml). Overweeg "&FIXED(MAX({S_CONC1},{S_CONC2}),0,TRUE)&" mg/ml of een groter volume.","")',
    '=IF(AND(D19="Anders (handmatig %)",NOT(ISNUMBER(D20))),"Vul bij \'Anders\' een handmatig percentage in.","")',
]
# Verborgen hulpkolommen L (tekst) en M (volgnummer van niet-lege teksten); de zichtbare
# regels tonen alleen de echte waarschuwingen, bovenaan, zonder lege tussenrijen.
W0 = CW + 2
W1 = W0 + len(warns) - 1
for i, f in enumerate(warns):
    r = W0 + i
    ws[f"L{r}"] = f
    ws[f"M{r}"] = f'=IF(L{r}<>"",COUNTIF(L${W0}:L{r},"?*"),"")'
ws.column_dimensions["L"].hidden = True
ws.column_dimensions["M"].hidden = True
AANTAL = f'COUNTIF($L${W0}:$L${W1},"?*")'
for k in range(1, len(warns) + 1):
    r = W0 + k - 1
    ws.merge_cells(f"C{r}:I{r}")
    tekst = f'IFERROR(INDEX($L${W0}:$L${W1},MATCH({k},$M${W0}:$M${W1},0)),"")'
    if k == 1:
        ws[f"C{r}"] = f'=IF({AANTAL}=0,"Geen bijzonderheden gevonden.",{tekst})'
        ws[f"B{r}"] = f'=IF({AANTAL}=0,"✓","⚠")'
    else:
        ws[f"C{r}"] = f"={tekst}"
        ws[f"B{r}"] = f'=IF(LEN(C{r})>0,"⚠","")'
    ws[f"C{r}"].font = font(9.5, ORANJE, name=FS)
    ws[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="center", indent=1)
    ws[f"B{r}"].font = Font(name="Segoe UI Symbol", size=11, color=ORANJE)
    ws[f"B{r}"].alignment = midden
    ws.row_dimensions[r].height = 30
    ws.conditional_formatting.add(f"B{r}:I{r}", FormulaRule(formula=[f'$B${r}="⚠"'], fill=cf_fill(ORANJE_L)))
ws.conditional_formatting.add(f"B{W0}:I{W0}", FormulaRule(formula=[f'$B${W0}="✓"'], fill=cf_fill(GROEN_L),
                                                          font=Font(color=GROEN_D)))
EIND = W1
ws.freeze_panes = "A5"
ws.print_area = f"A1:J{EIND}"
ws.print_options.horizontalCentered = True
ws.page_setup.orientation = "portrait"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins.left = ws.page_margins.right = 0.4
ws.page_margins.top = ws.page_margins.bottom = 0.5
ws.oddFooter.left.text = "Opioïdrotatie - rekenhulp"
ws.oddFooter.left.size = 8
ws.oddFooter.center.text = "Pagina &P van &N"
ws.oddFooter.center.size = 8
ws.oddFooter.right.text = "&D &T"
ws.oddFooter.right.size = 8

# ================================================================== Omrekentabel
qs = wb.create_sheet("Omrekentabel")
cols = [n for n, *_ in drugs if n != "Methadon oraal"]
extra = ["Morfine pomp mg/uur", "Morfine pomp ml/uur"]
laatste = chr(ord("B") + len(cols) + len(extra))
kopbalk(qs, "Snelle omrekentabel", "Equianalgetisch, zonder reductie. Bij rotatie wegens bijwerkingen: 75% van deze waarden. "
        "Pompkolommen gebruiken de concentratie van tabblad Rotatie.", chr(ord(laatste) + 1))
qs.column_dimensions["A"].width = 2
qs.column_dimensions["B"].width = 16
for j in range(len(cols) + len(extra)):
    qs.column_dimensions[chr(ord("C") + j)].width = 12.5
tabelkop(qs, 5, 2, ["OME\n(mg morfine oraal / 24 u)"] + cols + extra, 46)
qs.cell(row=6, column=2, value="").fill = fill(LICHT)
for j, n in enumerate(cols):
    c = qs.cell(row=6, column=3 + j, value=f'=INDEX({UNITS},MATCH("{n}",{NAMES},0))')
    c.font = f_klein
    c.fill = fill(LICHT)
    c.alignment = midden
for j, t in enumerate(("mg/uur", "ml/uur")):
    c = qs.cell(row=6, column=3 + len(cols) + j, value=t)
    c.font = f_klein
    c.fill = fill(LICHT)
    c.alignment = midden
for i, ome in enumerate([30, 60, 90, 120, 180, 240, 300, 360, 480, 600]):
    r = 7 + i
    band = fill(GRIJS_L) if i % 2 else fill(WIT)
    invoer(qs.cell(row=r, column=2), ome)
    for j, n in enumerate(cols):
        c = qs.cell(row=r, column=3 + j, value=f'=$B{r}/INDEX({FACT},MATCH("{n}",{NAMES},0))')
        c.number_format = "0.0"
        c.font = f_waarde
        c.fill = band
        c.alignment = midden
        c.border = rand_cel
    m = qs.cell(row=r, column=3 + len(cols), value=f"=$B{r}/{S_RATIO}/24")
    m.number_format = "0.00"
    l_ = qs.cell(row=r, column=4 + len(cols), value=f"=IF(N(Rotatie!$D$26)>0,$B{r}/{S_RATIO}/24/Rotatie!$D$26,0)")
    l_.number_format = "0.000"
    for c in (m, l_):
        c.font = Font(name=FS, size=10, color=GROEN_D)
        c.fill = fill(GROEN_L)
        c.alignment = midden
        c.border = rand_cel
    qs.row_dimensions[r].height = 20
qs["B18"] = "Witte OME-waarden in kolom B zijn aanpasbaar."
qs["B18"].font = f_klein
qs.freeze_panes = "C7"
qs.sheet_properties.tabColor = BLAUW_M
qs.page_setup.orientation = "landscape"
qs.page_setup.fitToWidth = 1
qs.page_setup.fitToHeight = 1
qs.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================== Doseercodes
ds = wb.create_sheet("Doseercodes", 2)
for col, w in zip("ABCDEFG", [2, 14, 20, 15, 20, 72, 2]):
    ds.column_dimensions[col].width = w
kopbalk(ds, "Eigen doseercodes", "Pharmacom-doseercodes die de rekenhulp niet zelf herkent. Een code hier gaat altijd vóór de "
        "automatische vertaling.", "G")
uitleg = [
    "De rekenhulp vertaalt doseercodes volgens de Centrale doseertabel (Health Base): frequentie - tijdsaanduiding - "
    "hoeveelheid - toedieningsvorm, eventueel 'M' + getal = niet meer dan ... per dag (bijv. P1CM6 = bij pijn 1 capsule, max. 6 per dag).",
    "Eenheden per 24 u = aantal tabletten/capsules/ml per dag (bij een pleister: aantal tegelijk geplakte pleisters). "
    "Zo nodig = J voor zo nodig/bij pijn: dan wordt het maximum per 24 u ingevuld; verlaag zelf naar het werkelijke gebruik.",
]
for i, t in enumerate(uitleg):
    c = ds.cell(row=5 + i, column=2, value=t)
    c.font = font(9, GRIJS)
    c.alignment = boven
    ds.merge_cells(start_row=5 + i, start_column=2, end_row=5 + i, end_column=6)
    ds.row_dimensions[5 + i].height = 30
tabelkop(ds, 8, 2, ["Code", "Eenheden per 24 u", "Zo nodig (J/N)", "Max. eenheden per 24 u", "Omschrijving"], 30)
for r in range(9, 59):
    for j in range(5):
        invoer(ds.cell(row=r, column=2 + j))
    ds.cell(row=r, column=6).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ds.row_dimensions[r].height = 20
ds["B9"], ds["C9"], ds["D9"], ds["F9"] = "VOORB1", 2, "N", "Voorbeeld (fictieve code, mag weg): 2 eenheden per 24 u, vast gebruik"
dv_jn = DataValidation(type="list", formula1='"J,N"', allow_blank=True)
ds.add_data_validation(dv_jn)
dv_jn.add("D9:D58")
ds.freeze_panes = "A9"
ds.sheet_properties.tabColor = GROEN
ds.page_setup.fitToWidth = 1
ds.page_setup.fitToHeight = 0
ds.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================== Toelichting
ts = wb.create_sheet("Toelichting")
ts.column_dimensions["A"].width = 2
ts.column_dimensions["B"].width = 118
ts.column_dimensions["C"].width = 2
kopbalk(ts, "Toelichting", "Werkwijze, rekenmethode, bronnen en verantwoording", "C")
ts.sheet_properties.tabColor = GRIJS
ts.page_setup.fitToWidth = 1
ts.page_setup.fitToHeight = 0
ts.sheet_properties.pageSetUpPr.fitToPage = True
H = "kop"
lines = [
    (H, "Werkwijze"),
    ("", "1.  Uitlezen: open in Pharmacom het dossier van de patiënt en druk op Ctrl+Alt+F3 (ook vanuit Excel). Pharmacom wordt "
         "dan zelf gekozen, zo nodig wordt de medicatiestatus (F3) geopend, en stap 1 wordt ingevuld. Handmatig invullen kan ook: "
         "kies per regel het middel en vul de totale dosis van de afgelopen 24 uur in (pleisters en continue fentanyl in µg/uur)."),
    ("", "2.  Kies de reden van rotatie (standaard: onvoldoende pijnstilling → 100%; bijwerkingen → 75%) en eventueel een extra reductie."),
    ("", "3.  Kies de concentratie morfine (10 of 1 mg/ml, eigen voorraad) en vul de kleinste instelstap van de pomp in. "
         "De berekening laat ook de pompstand bij de andere concentratie zien; bij een lage dosis (te grove afronding) of een "
         "spuit die binnen 24 uur leeg is, stelt de rekenhulp de andere concentratie voor."),
    ("", "4.  Lees de pompstand, dagdosis, doorbraakbolus en looptijd af in de tegels; de berekening staat eronder."),
    ("", "5.  Lees altijd de controles en waarschuwingen. Gebruik daarna 'Wis invoer' en sla het bestand niet op met patiëntgegevens."),
    ("", ""),
    (H, "Uitlezen uit Pharmacom"),
    ("", "Op veel computers mag Excel zelf geen programma's starten (beveiliging tegen macrovirussen). Daarom start het uitlezen "
         "buiten Excel, via de sneltoets Ctrl+Alt+F3 of de snelkoppeling 'Lees opioiden uit Pharmacom' op het bureaublad. "
         "Het script leest via de Java Access Bridge (alleen lezen), herkent opioïden aan de ATC-code, de sterkte aan de "
         "etiketnaam en vertaalt de doseercode. De kolom 'Uitgelezen uit Pharmacom' laat per regel zien wat er gelezen en hoe het vertaald is."),
    ("", "Zo nodig / bij pijn: wordt als 'Doorbraak' gezet met het MAXIMUM per 24 u uit de doseercode (bijv. P1CM6 = max. 6 per dag). "
         "Verlaag dit naar het werkelijke gebruik als dat minder is; staat er geen maximum in de code, dan blijft de dosis leeg. "
         "Oranje = controleren. Niet automatisch: fentanyl neusspray/zuigtablet, buprenorfine sublinguaal, onbekende codes."),
    ("", "Bovenaan staat welk dossier is uitgelezen - controleer of dit de juiste patiënt is. Pharmacom_naar_Excel.ps1 en "
         "Pharmacom_medicatiestatus.ps1 moeten in dezelfde map staan als deze werkmap."),
    ("", ""),
    (H, "Rekenmethode"),
    ("", "OME (mg morfine oraal/24 u) = Σ dosis × factor  →  morfine s.c./i.v. = OME ÷ verhouding oraal : s.c. (standaard 3)  →  "
         "× % volgens reden × (1 − extra reductie)  →  ÷ 24 = mg/uur  →  ÷ concentratie = ml/uur."),
    ("", "Voorbeeld (ingevuld): oxycodon oraal 80 + 30 mg = 110 mg × 1,5 = 165 mg OME → ÷ 3 = 55 mg morfine s.c./24 u bij "
         "onvoldoende pijnstilling (100%) = 2,29 mg/uur → bij 10 mg/ml: 0,23 ml/uur. Bij bijwerkingen (75%): 41,3 mg/24 u = 0,17 ml/uur."),
    ("", ""),
    (H, "Kleuren"),
    ("", "Wit vlak met blauwe rand = invulveld of aanpasbare aanname.  Lichtgrijs = berekend.  Groen = belangrijkste uitkomst.  "
         "Oranje = controleren of waarschuwing."),
    ("", ""),
    (H, "Bronnen"),
    ("", "• Hoofdbron: KNMP Informatorium Medicamentorum, inleidende tekst Analgetica (G806), rubriek Dosering - Opioïdrotatie, "
         "conversiefactoren, opioïdrotatietabel en Doorbraakpijn (geraadpleegd 14-10-2026). Deze factoren zijn afkomstig uit de "
         "NVA-richtlijn Pijn bij patiënten met kanker; de NHG-Standaard Pijn hanteert dezelfde omrekening."),
    ("", "• KNMP: rotatie wegens bijwerkingen → start 75% van de equivalente 24-uursdosis; wegens onvoldoende effect → equivalente dosis. "
         "Doorbraak: 1/6 van de dagdosis, max. 4-6× per dag; > 3× per dag nodig → onderhoudsdosis met 50% verhogen."),
    ("", "• Aanvullend: NVA-richtlijn Pijn bij COPD/hartfalen, module Opioïdrotatie (2019) - max. 140 µg/uur buprenorfine, max. tapentadol; "
         "Pallialine-pijnkaart - doorbraak 10-15% (lagere variant)."),
    ("", "• Doseercodes: Centrale doseertabel, Opbouw doseercodes (Health Base, helpkaart Pharmacom/Medicom)."),
    ("", "• Codeïne staat niet in de KNMP-rotatietabel (literatuurwaarde 1 : 10)."),
    ("", ""),
    (H, "Verantwoording"),
    ("", "Equianalgetische tabellen zijn benaderingen met beperkt wetenschappelijk bewijs; incomplete kruistolerantie, nier- en leverfunctie, "
         "leeftijd en comedicatie beïnvloeden de juiste dosis. Deze rekenhulp ondersteunt maar vervangt geen klinisch oordeel, lokaal "
         "protocol of dubbele controle. Bij twijfel: consultatieteam palliatieve zorg of pijnteam."),
    ("", ""),
    ("klein", VERSIE + "  ·  factoren volgens KNMP IM  ·  uitlezen uit Pharmacom  ·  github.com/ronaldsteenkamp/opioidrotatie"),
]
for i, (soort, t) in enumerate(lines):
    r = 5 + i
    c = ts.cell(row=r, column=2, value=t)
    if soort == H:
        c.font = Font(name=FS, size=12, color=BLAUW)
        c.border = Border(bottom=side(GROEN, "medium"))
        ts.row_dimensions[r].height = 24
    elif soort == "klein":
        c.font = font(8, BLAUW_M, italic=True)
    else:
        c.font = font(10, GRIJS)
        regels = max(1, len(t) // 130 + 1)
        ts.row_dimensions[r].height = 16 * regels + 4 if t else 8
    c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)

# Volgorde tabbladen: Rotatie, Omrekenfactoren, Doseercodes, Omrekentabel, Toelichting
wb._sheets = [ws, fs, ds, qs, ts]
wb.active = 0

# Samenvoegen wist de opmaak van de cellen achter de eerste cel; zet de vulling terug,
# anders ontstaan er naden in gekleurde vlakken (tegels, patiëntregel).
for sh in wb.worksheets:
    for mr in sh.merged_cells.ranges:
        anker = sh.cell(mr.min_row, mr.min_col)
        if anker.fill is not None and anker.fill.fill_type:
            for rij in sh.iter_rows(min_row=mr.min_row, max_row=mr.max_row, min_col=mr.min_col, max_col=mr.max_col):
                for c in rij:
                    c.fill = copy(anker.fill)

# Overal Segoe UI (ook cellen zonder eigen opmaak)
for sh in wb.worksheets:
    for row in sh.iter_rows():
        for c in row:
            if c.value is not None and c.font is not None and c.font.name not in (F, FS, "Segoe UI Symbol"):
                c.font = Font(name=F, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic,
                              color=c.font.color)

wb.save(OUT)
print("saved", OUT)
