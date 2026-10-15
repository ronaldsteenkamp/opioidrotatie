"""pytest: onafhankelijke referentie tegen de werkmap, en de parser-port tegen de echte VBA.

- Tests zonder werkmap_uitkomsten.json draaien overal (Linux/CI): referentie en parser-port.
- Tests met werkmap_uitkomsten.json vergelijken met wat Excel echt uitrekende (maak die met
  excel_harness.ps1 op Windows; in CI als artefact van de Windows-job).
- Bekende bevindingen staan als xfail(strict=True): zodra een fout is opgelost, faalt de test
  en moet de markering weg (zo blijft het reviewrapport en de code in de pas).
"""
import json
import math
import pathlib
import re

import pytest

import casussen as C
import parser_port as P
import referentie as R

HIER = pathlib.Path(__file__).parent
UIT = HIER / "werkmap_uitkomsten.json"
WERKMAP = json.loads(UIT.read_text(encoding="utf-8-sig")) if UIT.exists() else None
nodig_werkmap = pytest.mark.skipif(WERKMAP is None, reason="werkmap_uitkomsten.json ontbreekt (draai excel_harness.ps1)")


def _num(x):
    return x if isinstance(x, (int, float)) else None


def _per_id(soort):
    return {r["id"]: r for r in WERKMAP[soort]} if WERKMAP else {}


# ================================================================ referentie: README-voorbeelden
@pytest.mark.parametrize("middel, dosis, ome", [
    ("Morfine s.c./i.v.", 20, 60), ("Oxycodon oraal", 40, 60), ("Oxycodon s.c./i.v.", 20, 60),
    ("Hydromorfon oraal", 12, 60), ("Hydromorfon s.c./i.v.", 4, 60), ("Fentanyl pleister", 25, 60),
    ("Buprenorfine pleister", 25, 60), ("Tramadol oraal", 300, 60), ("Tapentadol oraal", 150, 60),
    ("Fentanyl oromucosaal", 1200, 60), ("Fentanyl nasaal", 200, 60), ("Methadon oraal", 15, 60),
])
def test_readme_voorbeelden(middel, dosis, ome):
    assert R.ome_regel(middel, dosis) == pytest.approx(ome)


@pytest.mark.parametrize("dosis, ome", [(15, 60), (20, 80), (21, 126), (50, 300), (51, 408), (20.5, 123)])
def test_methadon_staffel(dosis, ome):
    assert R.ome_regel("Methadon oraal", dosis) == pytest.approx(ome)


@pytest.mark.parametrize("ome, mg", [(60, 15), (80, 20), (100, 20), (120, 20), (121, 121 / 6), (300, 50), (350, 50),
                                     (400, 50), (500, 62.5)])
def test_methadon_omgekeerd(ome, mg):
    assert R.methadon_naar_dosis(ome) == pytest.approx(mg)


def test_methadon_omgekeerd_nooit_boven_ome():
    """Terugrekenen naar methadon mag nooit een dosis geven die (via de staffel) méér OME is."""
    for ome in range(1, 1000):
        d = R.methadon_naar_dosis(ome)
        assert R.ome_regel("Methadon oraal", d) <= ome + 1e-9


def test_buprenorfine_interne_consistentie():
    """Werkmap-opmerking: 'KNMP-tabel: 52,5 µg/uur = 120 mg morfine oraal'; factor 2,4 geeft 126 (bevinding L-01)."""
    assert R.ome_regel("Buprenorfine pleister", 52.5) == pytest.approx(126)


# ================================================================ pomp: werkmap == referentie
@nodig_werkmap
@pytest.mark.parametrize("cas", [c for c in C.POMP if all(isinstance(d, (int, float)) for _, d in c["regels"])],
                         ids=lambda c: c["id"])
def test_pomp_werkmap_gelijk_aan_referentie(cas):
    w = _per_id("pomp")[cas["id"]]
    egfr = cas["egfr"] if isinstance(cas["egfr"], (int, float)) else None
    ref = R.pomp([tuple(r) for r in cas["regels"]], cas["reden"], cas["extra"], cas["handmatig"], egfr,
                 cas["conc"], cas["stap"], cas["volume"])
    paren = [("D41", ref.ome), ("D44", ref.dagdosis), ("D45", ref.mg_uur), ("D46", ref.ml_exact), ("D47", ref.ml),
             ("D48", ref.echt), ("D49", ref.afwijking), ("D51", ref.looptijd), ("D52", ref.bolus), ("D56", ref.alt_ml)]
    for cel, verwacht in paren:
        assert _num(w[cel]) == pytest.approx(verwacht, abs=1e-9), f"{cas['id']} {cel}"


@nodig_werkmap
@pytest.mark.parametrize("cas", [c for c in C.POMP if all(isinstance(d, (int, float)) for _, d in c["regels"])],
                         ids=lambda c: c["id"])
def test_pomp_waarschuwingen(cas):
    w = " ".join(_per_id("pomp")[cas["id"]]["waarschuwingen"])
    egfr = cas["egfr"] if isinstance(cas["egfr"], (int, float)) else None
    ref = R.pomp([tuple(r) for r in cas["regels"]], cas["reden"], cas["extra"], cas["handmatig"], egfr,
                 cas["conc"], cas["stap"], cas["volume"])
    tekst = {"methadon": "Methadon:", "snelwerkend fentanyl": "Snelwerkend fentanyl", "fentanylpleister": "Fentanylpleister",
             "fentanyl parenteraal": "Fentanyl s.c./i.v.", "buprenorfine": "Buprenorfinepleister", "egfr<30": "eGFR < 30",
             "hoge dosis": "Hoge dosis", "boven max": "boven de maximale dagdosis", "afronding": "Afronding op de pompstap",
             "spuit <24u": "binnen 24 uur leeg", "volledige dosis": "Volledige equianalgetische dosis"}
    for code in ref.waarschuwingen:
        assert tekst[code] in w, f"{cas['id']}: waarschuwing '{code}' ontbreekt"


# ---------------------------------------------------------------- bekende bevindingen (xfail strict)
@nodig_werkmap
@pytest.mark.parametrize("id", ["P11", "P13", "P17"])
def test_pompstand_nooit_boven_berekende_dosis(id):   # K-02 hersteld
    assert _per_id("pomp")[id]["D49"] <= 1e-9


@nodig_werkmap
@pytest.mark.xfail(strict=True, reason="H-03: eGFR als tekst ('<15') geeft geen nierwaarschuwing")
def test_egfr_tekst_kleiner_dan_15_waarschuwt():
    assert any("eGFR < 30" in t for t in _per_id("pomp")["P24"]["waarschuwingen"])


@nodig_werkmap
@pytest.mark.xfail(strict=True, reason="M-01: negatieve dosis verlaagt de OME zonder melding")
def test_negatieve_dosis_geeft_melding():
    w = _per_id("pomp")["P26"]
    assert any("✓" not in c for c in w["controle"]) or w["G16"] != 60


@nodig_werkmap
@pytest.mark.parametrize("id", ["P12", "P14"])
def test_geen_pompstand_nul_in_tegel(id):
    assert _per_id("pomp")[id]["tegels"][0] not in ("0,00 ml/uur", "0.00 ml/uur")


# ================================================================ omzetten
@nodig_werkmap
@pytest.mark.parametrize("cas", C.OMZET, ids=lambda c: c["id"])
def test_omzet_dosis(cas):
    w = _per_id("omzet")[cas["id"]]
    ome = R.ome_totaal([tuple(r) for r in cas["regels"]])
    toeg = ome * R.PCT[cas["reden"]]
    assert w["C14"] == pytest.approx(ome)
    assert w["C22"] == pytest.approx(R.omzet_dosis(toeg, cas["doel"]))


def _advies_afwijking(tekst):
    m = re.search(r"\(([+-]\d+)%\)\s*$", tekst)
    return int(m.group(1)) / 100 if m else None


@nodig_werkmap
@pytest.mark.xfail(strict=True, reason="H-01: advies onderhoud kan >10% boven de berekende dosis liggen (Z02 +28%, Z13 +44%)")
@pytest.mark.parametrize("id", ["Z02", "Z13"])
def test_advies_onderhoud_niet_meer_dan_10pct_boven(id):
    w = _per_id("omzet")[id]
    groen = w["C26"] if _advies_afwijking(w["C26"]) is not None else w["C25"]
    assert _advies_afwijking(groen) <= 0.10


@nodig_werkmap
@pytest.mark.parametrize("id", ["Z02", "Z11", "Z12"])
def test_advies_doorbraak_nooit_boven_berekend(id):
    w = _per_id("omzet")[id]
    afw = _advies_afwijking(w["C43"].replace(" mg per keer", ""))
    assert afw is None or afw <= 0     # K-01 hersteld: naar beneden afgerond


@nodig_werkmap
@pytest.mark.xfail(strict=True, reason="H-02: advies boven max. dagdosis wordt toch als preparaat + doseercode getoond")
def test_geen_advies_boven_max_dagdosis():
    assert "TRAMADOL" not in _per_id("omzet")["Z09"]["C25"]


@pytest.mark.parametrize("doel, dosis", [("Fentanyl pleister", 18.75), ("Oxycodon oraal", 40), ("Hydromorfon oraal", 21)])
def test_referentie_veilige_keuze_nooit_erboven(doel, dosis):
    sterktes, f = C.STERKTES[doel]
    k = R.beste_preparaat(dosis, sterktes, f, beleid="veilig")
    assert k.afwijking <= 1e-9


# ================================================================ parser: port == VBA
def _vba_code(tekst):
    if tekst == "niet herkend":
        return None
    nl = lambda s: float(s.replace(",", "."))  # noqa: E731
    m = re.search(r"\| min (\S+) \| max (\S+) \| zo nodig (\S+) \| aantal (\S+) \| vorm (\S*)$", tekst)
    return dict(min=nl(m.group(1)), max=nl(m.group(2)), zo=m.group(3) == "Waar", aantal=nl(m.group(4)), vorm=m.group(5))


@nodig_werkmap
@pytest.mark.parametrize("rij", WERKMAP["codes"] if WERKMAP else [], ids=lambda r: r["code"])
def test_port_vertaalcode_gelijk_aan_vba(rij):
    vba = _vba_code(rij["vba"])
    d = P.vertaal_code(rij["code"])
    if vba is None:
        assert not d.herkend
        return
    assert d.herkend
    assert d.per_dag_min == pytest.approx(vba["min"], abs=1e-9)
    assert d.per_dag_max == pytest.approx(vba["max"], abs=1e-9)
    assert d.zo_nodig == vba["zo"]
    assert d.aantal == pytest.approx(vba["aantal"])
    assert d.vorm == vba["vorm"]


@nodig_werkmap
@pytest.mark.parametrize("rij", WERKMAP["sterktes"] if WERKMAP else [], ids=lambda r: r["naam"] + "|" + r["atc"])
def test_port_sterkte_gelijk_aan_vba(rij):
    stof = P.opioid(rij["atc"], rij["naam"])
    if rij["vba"] == "geen opioid":
        assert stof == ""
        return
    delen = [d.strip() for d in rij["vba"].split("|")]
    assert stof == delen[0]
    weg = P.toedieningsweg(rij["naam"])
    assert weg == delen[1]
    assert P.tabel_naam(stof, weg) == delen[2]
    s = P.lees_sterkte(rij["naam"], stof)
    if delen[3] == "sterkte?":
        assert not s.gevonden
    else:
        waarde = float(delen[3].split()[0].replace(",", "."))
        assert s.waarde == pytest.approx(waarde)


# ---------------------------------------------------------------- parser: bedoelde betekenis
def _per_dag(d):
    return d.per_dag_max if d.zo_nodig else d.per_dag_min


@pytest.mark.parametrize("code, bedoeld", [(k, v) for k, v in C.CODES.items() if v is not None and not isinstance(v, tuple)])
def test_code_vaste_dosering(code, bedoeld):
    d = P.vertaal_code(code)
    assert d.herkend and not d.zo_nodig
    assert d.per_dag_min == pytest.approx(bedoeld)


@pytest.mark.parametrize("code, bedoeld", [(k, v) for k, v in C.CODES.items() if isinstance(v, tuple) and v[0] == "zo nodig"])
def test_code_zo_nodig(code, bedoeld):
    d = P.vertaal_code(code)
    assert d.herkend and d.zo_nodig
    assert (_per_dag(d) if d.per_dag_max > 0 else None) == bedoeld[1]


@pytest.mark.parametrize("naam, atc, middel, sterkte", [s for s in C.STERKTES_NAMEN if s[2] is not None],
                         ids=lambda x: str(x))
def test_naam_naar_middel(naam, atc, middel, sterkte):
    stof = P.opioid(atc, naam)
    gevonden = P.tabel_naam(stof, P.toedieningsweg(naam)) if stof else ""
    bekend_fout = {("HYDROMORFON CAPSULE MGA 4MG", ""): "H-04", ("PALLADON SLOW 8MG", ""): "M-04",
                   ("FENTANYL SPRAY NASAAL 100UG/DOSIS", "N02AB03"): "M-03", ("TRANSTEC 35 MICROGRAM/UUR", "N02AE01"): "M-04"}
    if (naam, atc) in bekend_fout:
        pytest.xfail(f"bevinding {bekend_fout[(naam, atc)]}")
    assert gevonden == middel
    s = P.lees_sterkte(naam, stof)
    if naam != "OXYCONTIN 10":
        per = s.waarde / s.per_aantal if s.per == "ML" and s.per_aantal > 0 else s.waarde
        assert per == pytest.approx(sterkte)


# ================================================================ layout: VBA-constanten == werkmap_opbouw.py
@pytest.fixture(scope="module")
def basis(tmp_path_factory):
    """Bouwt de basiswerkmap met openpyxl (kan op Linux) en geeft (werkmap, Pharmacom.bas-tekst)."""
    openpyxl = pytest.importorskip("openpyxl")
    import subprocess
    import sys
    bron = HIER.parent / "bron"
    uit = tmp_path_factory.mktemp("bouw") / "basis.xlsx"
    subprocess.run([sys.executable, "-I", str(bron / "werkmap_opbouw.py"), str(uit)], check=True, capture_output=True)
    return openpyxl.load_workbook(uit), (bron / "Pharmacom.bas").read_text(encoding="utf-8")


def _const(bas, naam):
    return re.search(rf"Const {naam} As \w+ = \"?([^\"\s']+)", bas).group(1)


def test_layout_blok1_en_egfr(basis):
    wb, bas = basis
    ws = wb["Rotatie"]
    assert ws.cell(9, 3).value.startswith("Middel")                  # kop blok 1 staat op rij R0-1
    r0, r1 = int(_const(bas, "R_EERST")), int(_const(bas, "R_LAATST"))
    assert ws[f"C{r0 - 1}"].value.startswith("Middel") and ws[f"B{r1 + 1}"].value.startswith("Totaal")
    egfr = _const(bas, "CEL_EGFR")
    assert "eGFR" in ws[f"B{egfr[1:]}"].value


def test_layout_doseercodes(basis):
    wb, bas = basis
    ds = wb["Doseercodes"]
    m = re.search(r'ws\.Range\("([A-Z])(\d+):([A-Z])(\d+)"\)\.Cells', bas)
    kol, rij = m.group(1), int(m.group(2))
    assert ds[f"{kol}{rij - 1}"].value == "Code"                     # kopregel direct boven het leesbereik


# ================================================================ end-to-end Verwerk en eigen codes
@nodig_werkmap
def test_eigen_doseercode_wordt_gebruikt():
    assert "eigen tabel" in WERKMAP["extra"]["eigen_code_volgens_tabblad"]


@nodig_werkmap
@pytest.mark.xfail(strict=True, reason="H-04: hydromorfon zonder ATC wordt als morfine ingevuld, niet oranje")
def test_hydromorfon_zonder_atc():
    regel = {r["id"]: r for r in WERKMAP["verwerk"]}["V02"]["regels"][0]
    assert regel["middel"] == "Hydromorfon oraal" or regel["oranje"]


@nodig_werkmap
@pytest.mark.xfail(strict=True, reason="M-05: eGFR met tekst erachter wordt een onzinnig getal (452009)")
def test_egfr_met_toelichting():
    assert {r["id"]: r for r in WERKMAP["verwerk"]}["V07"]["D22"] in (45, None)


@nodig_werkmap
def test_verwerk_normaal():
    v = {r["id"]: r for r in WERKMAP["verwerk"]}["V01"]
    assert [(r["middel"], r["dosis"], r["soort"]) for r in v["regels"]] == [
        ("Oxycodon oraal", 40, "Onderhoud"), ("Oxycodon oraal", 30, "Doorbraak"), ("Fentanyl pleister", 12, "Onderhoud")]
    assert v["D22"] == 45
