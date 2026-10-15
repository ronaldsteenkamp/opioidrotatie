"""Testcasussen (verzonnen, geen patiëntgegevens). `python casussen.py` schrijft casussen.json voor
excel_harness.ps1, dat ze door de echte werkmap haalt."""
import json
import pathlib

O = "Onvoldoende pijnstilling"
B = "Bijwerkingen"


def c(id, oms, regels, reden=O, extra=0.0, egfr=None, conc=10, stap=0.01, volume=50, handmatig=None):
    return dict(id=id, oms=oms, regels=regels, reden=reden, extra=extra, egfr=egfr, conc=conc, stap=stap,
                volume=volume, handmatig=handmatig)


POMP = [
    c("P01", "60 mg morfine oraal, onvoldoende effect", [["Morfine oraal", 60]]),
    c("P02", "fentanylpleister 25 µg/uur, bijwerkingen (75%)", [["Fentanyl pleister", 25]], reden=B),
    c("P03", "oxycodon 2x20 mg + oxycodon zo nodig P1CM6 5 mg (max 6 = 30 mg)",
      [["Oxycodon oraal", 40], ["Oxycodon oraal", 30]]),
    c("P04", "combinatie morfine 60 + tramadol 200 + fentanylpleister 12",
      [["Morfine oraal", 60], ["Tramadol oraal", 200], ["Fentanyl pleister", 12]]),
    c("P05", "methadon 15 mg", [["Methadon oraal", 15]]),
    c("P06", "methadon 20 mg", [["Methadon oraal", 20]]),
    c("P07", "methadon 21 mg", [["Methadon oraal", 21]]),
    c("P08", "methadon 50 mg", [["Methadon oraal", 50]]),
    c("P09", "methadon 51 mg", [["Methadon oraal", 51]]),
    c("P10", "methadon 20,5 mg (tussen de treden)", [["Methadon oraal", 20.5]]),
    c("P11", "zeer laag: 5 mg morfine oraal, 10 mg/ml, stap 0,01", [["Morfine oraal", 5]]),
    c("P12", "zeer laag: 5 mg morfine oraal, 10 mg/ml, stap 0,1", [["Morfine oraal", 5]], stap=0.1),
    c("P13", "zeer laag: 5 mg morfine oraal, 1 mg/ml, stap 0,1", [["Morfine oraal", 5]], conc=1, stap=0.1),
    c("P14", "zeer laag: 2 mg morfine oraal (afronding naar 0)", [["Morfine oraal", 2]]),
    c("P15", "zeer hoog: oxycodon 800 mg (1200 OME), spuit 20 ml", [["Oxycodon oraal", 800]], volume=20),
    c("P16", "morfine 90 mg, 10 mg/ml, stap 0,1", [["Morfine oraal", 90]], stap=0.1),
    c("P17", "morfine 90 mg, 10 mg/ml, stap 0,05 (precies op de helft)", [["Morfine oraal", 90]], stap=0.05),
    c("P18", "morfine 90 mg, 1 mg/ml, stap 0,05, spuit 20", [["Morfine oraal", 90]], conc=1, stap=0.05, volume=20),
    c("P19", "morfine 90 mg, 1 mg/ml, stap 0,1, spuit 50", [["Morfine oraal", 90]], conc=1, stap=0.1),
    c("P20", "eGFR 29", [["Morfine oraal", 60]], egfr=29),
    c("P21", "eGFR 30", [["Morfine oraal", 60]], egfr=30),
    c("P22", "eGFR ontbrekend", [["Morfine oraal", 60]], egfr=None),
    c("P23", "eGFR '>90' (tekst, handmatig)", [["Morfine oraal", 60]], egfr=">90"),
    c("P24", "eGFR '<15' (tekst, handmatig getypt)", [["Morfine oraal", 60]], egfr="<15"),
    c("P25", "lege invoer", []),
    c("P26", "negatieve dosis (geplakt, langs de validatie)", [["Morfine oraal", 120], ["Oxycodon oraal", -40]]),
    c("P27", "tekst in getalveld '60 mg'", [["Morfine oraal", "60 mg"]]),
    c("P28", "getal als tekst '60' (geplakt)", [["Morfine oraal", "60"]]),
    c("P29", "tramadol 500 mg (boven max 400)", [["Tramadol oraal", 500]]),
    c("P30", "hoge dosis precies 240 OME", [["Morfine oraal", 240]]),
    c("P31", "bijwerkingen + extra reductie 25%", [["Morfine oraal", 120]], reden=B, extra=0.25),
    c("P32", "Anders (handmatig %) zonder percentage", [["Morfine oraal", 60]], reden="Anders (handmatig %)"),
    c("P33", "fentanyl s.c./i.v. 25 µg/uur", [["Fentanyl s.c./i.v. continu", 25]]),
    c("P34", "fentanyl oromucosaal 1200 µg/24 u", [["Fentanyl oromucosaal", 1200]]),
    c("P35", "buprenorfinepleister 52,5 µg/uur", [["Buprenorfine pleister", 52.5]]),
]

OMZET = [
    dict(id="Z01", oms="165 OME (oxy 80+30) -> fentanylpleister", regels=[["Oxycodon oraal", 80], ["Oxycodon oraal", 30]],
         reden=O, doel="Fentanyl pleister", freq=None),
    dict(id="Z02", oms="60 OME, bijwerkingen (45) -> fentanylpleister", regels=[["Morfine oraal", 60]], reden=B,
         doel="Fentanyl pleister", freq=None),
    dict(id="Z03", oms="60 OME -> oxycodon oraal", regels=[["Morfine oraal", 60]], reden=O, doel="Oxycodon oraal", freq=None),
    dict(id="Z04", oms="105 OME -> hydromorfon oraal", regels=[["Oxycodon oraal", 70]], reden=O,
         doel="Hydromorfon oraal", freq=None),
    dict(id="Z05", oms="60 OME -> methadon", regels=[["Morfine oraal", 60]], reden=O, doel="Methadon oraal", freq=None),
    dict(id="Z06", oms="100 OME -> methadon (gat in staffel)", regels=[["Morfine oraal", 100]], reden=O,
         doel="Methadon oraal", freq=None),
    dict(id="Z07", oms="350 OME -> methadon (gat in staffel)", regels=[["Morfine oraal", 350]], reden=O,
         doel="Methadon oraal", freq=None),
    dict(id="Z08", oms="500 OME -> methadon", regels=[["Morfine oraal", 500]], reden=O, doel="Methadon oraal", freq=None),
    dict(id="Z09", oms="300 OME -> tramadol (boven max)", regels=[["Morfine oraal", 300]], reden=O, doel="Tramadol oraal",
         freq=None),
    dict(id="Z10", oms="45 mg morfine -> oxycodon 3x per dag", regels=[["Morfine oraal", 45]], reden=O,
         doel="Oxycodon oraal", freq=3),
    dict(id="Z11", oms="40 OME bijwerkingen (30) -> fentanylpleister", regels=[["Morfine oraal", 40]], reden=B,
         doel="Fentanyl pleister", freq=None),
    dict(id="Z12", oms="90 OME -> buprenorfinepleister", regels=[["Morfine oraal", 90]], reden=O,
         doel="Buprenorfine pleister", freq=None),
    dict(id="Z13", oms="20 OME -> fentanylpleister (lager dan kleinste)", regels=[["Morfine oraal", 20]], reden=O,
         doel="Fentanyl pleister", freq=None),
    dict(id="Z14", oms="1300 OME -> fentanylpleister (zeer hoog)", regels=[["Oxycodon oraal", 866.67]], reden=O,
         doel="Fentanyl pleister", freq=None),
]

# Sterktes zoals op tabblad Preparaten (onderhoud), voor de onafhankelijke keuze
STERKTES = {
    "Morfine oraal": ([10, 15, 30, 60, 100], 2),
    "Oxycodon oraal": ([5, 10, 15, 20, 30, 40, 60, 80, 120], 2),
    "Hydromorfon oraal": ([4, 8, 16, 24], 2),
    "Fentanyl pleister": ([12, 25, 37, 50, 75, 100], None),
    "Buprenorfine pleister": ([5, 10, 15, 20, 35, 52.5, 70], None),
    "Tramadol oraal": ([50, 100, 150, 200], 2),
    "Tapentadol oraal": ([25, 50, 100, 150, 200, 250], 2),
    "Methadon oraal": ([2, 5, 10, 20, 40, 50], 2),
}

# ---------------------------------------------------------------- parser (Pharmacom.bas)
# Doseercodes met de BEDOELDE betekenis (eenheden per 24 u; zo nodig -> max per 24 u).
# None = de code hoort niet herkend te worden (oranje, zelf invullen) of de betekenis is onzeker.
CODES = {
    "2.1T": 2, "3.1T": 3, "6.1T": 6, "2.2T": 4, "3.12T": (3, 6), "2.H1T": (1, 2), "2.1C": 2, "2.1S": 2,
    "P1CM6": ("zo nodig", 6), "P2TM6": ("zo nodig", 6), "4P1TM4": ("zo nodig", 4), "P1C": ("zo nodig", None),
    "Z1T": ("zo nodig", None), "3Z1T": ("zo nodig", 3), "P5MLM30": ("zo nodig", 30), "2.1TM2": 2,
    "MA1T": 2, "M1A2T": 3, "MAN1T": 3, "A1M1T": 2,
    "D31PL": 1 / 3, "D32PL": 2 / 3, "D41PL": 1 / 4, "1W1PL": 1 / 7, "1.1PL": 1, "OD1T": 0.5, "ZO-1T": 1 / 7,
    "4.5ML": 20, "2D1T": 2, "3D1-2T": None, "1D1PL72U": None,
    "3U1T": None, "4U1T": None,          # betekenis 'U' onbekend: 3x per dag of elke 3 uur? (vraag)
    "12.1T": (1, 2), "10.1T": None, "Z1M4": ("zo nodig", 4), "Z12M8": ("zo nodig", 8),
}

# Etiketnamen: (naam, ATC, bedoeld middel op tabblad Omrekenfactoren, bedoelde sterkte per eenheid)
STERKTES_NAMEN = [
    ("OXYCODON TABLET MGA 20MG", "N02AA05", "Oxycodon oraal", 20),
    ("OXYCODON/NALOXON TABLET MGA 10/5MG", "N02AA55", "Oxycodon oraal", 10),
    ("TARGINACT TABLET MGA 20/10MG", "N02AA55", "Oxycodon oraal", 20),
    ("MORFINE DRANK 10MG/5ML", "N02AA01", "Morfine oraal", 2),
    ("MORFINE HCL DRANK 2MG/ML", "N02AA01", "Morfine oraal", 2),
    ("MORFINE INJVLST 10MG/ML", "N02AA01", "Morfine s.c./i.v.", 10),
    ("MORFINE ZETPIL 10MG", "N02AA01", "Morfine oraal", 10),
    ("FENTANYL PLEISTER 12UG/UUR", "N02AB03", "Fentanyl pleister", 12),
    ("FENTANYL PLEISTER 37,5UG/UUR", "N02AB03", "Fentanyl pleister", 37.5),
    ("FENTANYL MATRIXPLEISTER 25MICROGRAM/UUR", "N02AB03", "Fentanyl pleister", 25),
    ("DUROGESIC 50 PLEISTER 50UG/UUR", "N02AB03", "Fentanyl pleister", 50),
    ("FENTANYL 12 MICROGRAM/UUR PLEISTER", "N02AB03", "Fentanyl pleister", 12),
    ("BUPRENORFINE PLEISTER 52,5UG/UUR 3 DAGEN", "N02AE01", "Buprenorfine pleister", 52.5),
    ("TRANSTEC 35 MICROGRAM/UUR", "N02AE01", "Buprenorfine pleister", 35),
    ("BUPRENORFINE TABLET SUBLINGUAAL 0,4MG", "N07BC01", None, 0.4),
    ("ABSTRAL TABLET SUBLINGUAAL 100MICROGRAM", "N02AB03", "Fentanyl oromucosaal", 100),
    ("INSTANYL NEUSSPRAY 50UG/DOSIS", "N02AB03", "Fentanyl nasaal", 50),
    ("FENTANYL SPRAY NASAAL 100UG/DOSIS", "N02AB03", "Fentanyl nasaal", 100),
    ("FENTANYL INJVLST 50UG/ML", "N02AB03", "Fentanyl s.c./i.v. continu", 50),
    ("HYDROMORFON CAPSULE MGA 4MG", "N02AA03", "Hydromorfon oraal", 4),
    ("HYDROMORFON CAPSULE MGA 4MG", "", "Hydromorfon oraal", 4),
    ("PALLADON SLOW 8MG", "", "Hydromorfon oraal", 8),
    ("TRAMADOL/PARACETAMOL TABLET 37,5/325MG", "N02AJ13", "Tramadol oraal", 37.5),
    ("PARACETAMOL/CODEINE TABLET 500/20MG", "N02AJ06", "Codeïne oraal", 20),
    ("CODEINE TABLET 10MG", "R05DA04", "Codeïne oraal", 10),
    ("METHADON DRANK 1MG/ML", "N07BC02", "Methadon oraal", 1),
    ("TAPENTADOL TABLET MGA 50MG", "N02AX06", "Tapentadol oraal", 50),
    ("OXYCONTIN 10", "N02AA05", "Oxycodon oraal", 10),
    ("OXYCODON HCL INJVLST 10MG/ML", "N02AA05", "Oxycodon s.c./i.v.", 10),
    ("MORFINE TABLET MGA 30MG", "", "Morfine oraal", 30),
]

# End-to-end door Verwerk: [etiketnaam, doseercode, ATC, einddatum]
VERWERK = [
    dict(id="V01", oms="normaal: oxycodon MGA 20 2.1T + capsule 5 P1CM6 + fentanyl 12 D31PL",
         egfr=["45", "01-10-2026"],
         regels=[["OXYCODON TABLET MGA 20MG", "2.1T", "N02AA05", ""], ["OXYCODON CAPSULE 5MG", "P1CM6", "N02AA05", ""],
                 ["FENTANYL PLEISTER 12UG/UUR", "D31PL", "N02AB03", ""]]),
    dict(id="V02", oms="hydromorfon zonder ATC", egfr=None,
         regels=[["HYDROMORFON CAPSULE MGA 4MG", "2.1C", "", ""]]),
    dict(id="V03", oms="fentanylspray 'NASAAL'", egfr=None,
         regels=[["FENTANYL SPRAY NASAAL 100UG/DOSIS", "Z1M4", "N02AB03", ""]]),
    dict(id="V04", oms="code met U (elke 4 uur?)", egfr=None,
         regels=[["MORFINE TABLET 10MG", "4U1T", "N02AA01", ""]]),
    dict(id="V05", oms="einddatum met 2-cijferig jaar (31-12-30)", egfr=None,
         regels=[["OXYCODON TABLET MGA 20MG", "2.1T", "N02AA05", "31-12-30"]]),
    dict(id="V06", oms="eGFR '<15'", egfr=["<15", "01-10-2026"], regels=[["MORFINE TABLET MGA 30MG", "2.1T", "N02AA01", ""]]),
    dict(id="V07", oms="eGFR met tekst erachter", egfr=["45 (CKD-EPI 2009)", "01-10-2026"],
         regels=[["MORFINE TABLET MGA 30MG", "2.1T", "N02AA01", ""]]),
    dict(id="V08", oms="eGFR-datum met 2-cijferig jaar", egfr=["25", "01-10-26"],
         regels=[["MORFINE TABLET MGA 30MG", "2.1T", "N02AA01", ""]]),
    dict(id="V09", oms="eGFR decimaal '29,5'", egfr=["29,5", "01-10-2026"],
         regels=[["MORFINE TABLET MGA 30MG", "2.1T", "N02AA01", ""]]),
    dict(id="V10", oms="drank zonder /ML in de naam", egfr=None,
         regels=[["MORFINE DRANK 2MG", "4.5ML", "N02AA01", ""]]),
    dict(id="V11", oms="merknaam zonder ATC", egfr=None, regels=[["PALLADON SLOW 8MG", "2.1C", "", ""]]),
]

if __name__ == "__main__":
    p = pathlib.Path(__file__).with_name("casussen.json")
    p.write_text(json.dumps(dict(pomp=POMP, omzet=OMZET, codes=list(CODES), verwerk=VERWERK,
                                 sterktes=[[n, a] for n, a, *_ in STERKTES_NAMEN]), ensure_ascii=False, indent=1),
                 encoding="utf-8")
    print("geschreven:", p)
