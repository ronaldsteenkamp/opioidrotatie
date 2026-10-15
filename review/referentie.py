"""Onafhankelijke referentie-implementatie van de rekenkern van Opioidrotatie_morfinepomp.xlsm.

Geschreven vanuit de README (tabel Omrekenfactoren) en de functionele beschrijving, NIET door de
formules uit bron/werkmap_opbouw.py over te nemen. Zo vangt een vergelijking ook fouten die in
beide (code en README) tegelijk zitten niet, maar wel elke afwijking tussen bedoeling en werkmap.

Twee varianten voor de pompstand:
  - pomp(..., afronding="omlaag")   : standaard sinds herstel K-02 (nooit boven de berekende dosis)
  - pomp(..., afronding="excel")    : werkmap tot en met v1.7 (dichtstbijzijnde pompstap)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
import math

# ---------------------------------------------------------------- factoren uit de README
# Eenheid invoer -> factor naar orale morfine-equivalenten (OME, mg/24 u)
FACTOREN = {
    "Morfine oraal": 1.0,                       # referentie
    "Morfine s.c./i.v.": 3.0,                   # oraal : s.c. = 3 : 1
    "Oxycodon oraal": 1.5,                      # morfine : oxycodon = 1,5 : 1
    "Oxycodon s.c./i.v.": 1.5 * 2,              # oxycodon oraal : s.c. = 2 : 1
    "Hydromorfon oraal": 5.0,                   # 5 : 1
    "Hydromorfon s.c./i.v.": 5.0 * 3,           # hydromorfon oraal : s.c. = 3 : 1
    "Fentanyl pleister": 24 / 1000 * 100,       # µg/uur -> mg/24 u x 100
    "Fentanyl s.c./i.v. continu": 24 / 1000 * 100 * 3,   # parenteraal morfine : fentanyl 100 : 1, x 3 naar oraal
    "Fentanyl oromucosaal": 10 / 200,           # 200 µg = 10 mg morfine oraal (µg/24 u)
    "Fentanyl nasaal": 30 / 100,                # 100 µg = 30 mg morfine oraal (µg/24 u)
    "Buprenorfine pleister": 24 / 1000 * 100,   # 100 : 1
    "Tramadol oraal": 1 / 5,
    "Tapentadol oraal": 1 / 2.5,
    # Niet in de README-tabel (alleen in de werkmap); zie bevinding D-1:
    "Piritramide s.c./i.v.": 3 * 10 / 20,       # 20 mg piritramide = 10 mg morfine parenteraal
    "Codeïne oraal": 1 / 10,
}
MAX_DAG = {"Buprenorfine pleister": 140, "Tramadol oraal": 400, "Tapentadol oraal": 500, "Codeïne oraal": 240}

# Methadon (ZMC-equivalentietabel): 10 mg methadon = 40 (0-20 mg/24 u), 60 (21-50) of 80 (>50) mg morfine oraal
METHADON_STAFFEL = [(20, 4.0), (50, 6.0), (math.inf, 8.0)]

RATIO_ORAAL_SC = 3.0
PCT = {"Onvoldoende pijnstilling": 1.0, "Bijwerkingen": 0.75}
BOLUS_FRACTIE = 1 / 6
BOLUS_LAAG = 0.10
HOGE_DOSIS = 240
CONCENTRATIES = (10.0, 1.0)


def excel_round(x: float, n: int = 0) -> float:
    """ROUND() van Excel: half van nul af, na afronden op 15 significante cijfers."""
    if x == 0 or not math.isfinite(x):
        return x
    d = Decimal(repr(x))
    exp = d.adjusted()
    d = d.quantize(Decimal(1).scaleb(exp - 14), rounding=ROUND_HALF_UP)   # 15 significante cijfers
    return float(d.quantize(Decimal(1).scaleb(-n), rounding=ROUND_HALF_UP))


def methadon_factor(dosis: float) -> float:
    for tm, f in METHADON_STAFFEL:
        if dosis <= tm:
            return f
    raise ValueError(dosis)


def methadon_naar_dosis(ome: float) -> float:
    """Methadondosis d zodat d x factor(d) = OME; valt de OME in een gat van de staffel
    (80-120 of 300-400 mg), dan de hoogste dosis die er nog onder blijft."""
    ondergrens = 0.0
    for tm, f in METHADON_STAFFEL:
        d = ome / f
        if d <= ondergrens:            # gat in de staffel (of OME 0): blijf op de vorige grens
            return ondergrens if ondergrens > 0 else d
        if d <= tm:
            return d
        ondergrens = tm
    raise AssertionError("onbereikbaar: laatste trede is onbegrensd")


def ome_regel(middel: str, dosis: float) -> float:
    if middel == "Methadon oraal":
        return dosis * methadon_factor(dosis)
    return dosis * FACTOREN[middel]


def ome_totaal(regels) -> float:
    return sum(ome_regel(m, d) for m, d in regels)


@dataclass
class Pomp:
    ome: float
    sc_equi: float
    pct: float
    dagdosis: float
    mg_uur: float
    ml_exact: float
    ml: float
    echt: float
    afwijking: float
    ml_24u: float
    looptijd: float
    bolus: float
    bolus_laag: float
    alt_conc: float
    alt_ml: float
    waarschuwingen: list = field(default_factory=list)


def pomp(regels, reden="Onvoldoende pijnstilling", extra=0.0, handmatig=None, egfr=None,
         conc=10.0, stap=0.01, volume=50.0, afronding="omlaag") -> Pomp:
    ome = ome_totaal(regels)
    pct = (PCT[reden] if reden in PCT else (handmatig or 0.0)) * (1 - extra)
    sc = ome / RATIO_ORAAL_SC
    dag = sc * pct
    uur = dag / 24
    exact = uur / conc if conc > 0 else 0.0

    def rond(x, s):
        if s <= 0:
            return x
        if afronding == "omlaag":
            return math.floor(x / s + 1e-9) * s
        return excel_round(x / s) * s

    ml = rond(exact, stap)
    echt = ml * conc * 24
    afw = echt / dag - 1 if dag > 0 else 0.0
    loop = volume / ml if ml > 0 else 0.0
    alt = CONCENTRATIES[1] if conc == CONCENTRATIES[0] else CONCENTRATIES[0]
    alt_ml = rond(uur / alt, stap) if dag > 0 else 0.0
    p = Pomp(ome, sc, pct, dag, uur, exact, ml, echt, afw, ml * 24, loop, echt * BOLUS_FRACTIE,
             echt * BOLUS_LAAG, alt, alt_ml)
    w = p.waarschuwingen
    namen = [m for m, _ in regels]
    if any(m.startswith("Methadon") for m in namen):
        w.append("methadon")
    if "Fentanyl oromucosaal" in namen or "Fentanyl nasaal" in namen:
        w.append("snelwerkend fentanyl")
    if "Fentanyl pleister" in namen:
        w.append("fentanylpleister")
    if "Fentanyl s.c./i.v. continu" in namen:
        w.append("fentanyl parenteraal")
    if "Buprenorfine pleister" in namen:
        w.append("buprenorfine")
    if isinstance(egfr, (int, float)) and egfr < 30:
        w.append("egfr<30")
    if ome >= HOGE_DOSIS:
        w.append("hoge dosis")
    if any(m in MAX_DAG and d > MAX_DAG[m] for m, d in regels):
        w.append("boven max")
    if reden == "Onvoldoende pijnstilling" and extra == 0:
        w.append("volledige dosis")
    if abs(afw) > 0.1:
        w.append("afronding")
    if ome > 0 and 0 < loop < 24:
        w.append("spuit <24u")
    return p


# ---------------------------------------------------------------- omzetten naar een ander opioïd
def omzet_dosis(ome_toegepast: float, doel: str) -> float:
    if doel == "Methadon oraal":
        return methadon_naar_dosis(ome_toegepast)
    return ome_toegepast / FACTOREN[doel]


@dataclass
class Keuze:
    omschrijving: str
    dosis: float          # per 24 u (of µg/uur bij pleisters)
    afwijking: float      # t.o.v. berekend
    eenheden: int


def beste_preparaat(dosis: float, sterktes, keer_per_dag=None, max_per_keer=4, beleid="dichtst"):
    """Onafhankelijke keuze uit sterktes (één sterkte of twee sterktes per keer).
    keer_per_dag=None betekent pleister (dosis = sterkte x aantal).
    beleid='dichtst' : kleinste |afwijking| (wat de werkmap bedoelt)
    beleid='veilig'  : grootste dosis <= berekend (nooit erboven); valt terug op 'dichtst' als niets past."""
    f = keer_per_dag or 1
    opties = []
    for s in sterktes:
        for n in range(1, max_per_keer + 1):
            opties.append(Keuze(f"{n}x{s}", n * s * f, n * s * f / dosis - 1, n))
    for a in sterktes:
        for b in sterktes:
            if b < a:
                for n in range(1, max_per_keer):
                    opties.append(Keuze(f"{n}x{a}+1x{b}", (n * a + b) * f, (n * a + b) * f / dosis - 1, n + 1))
    if beleid == "veilig":
        onder = [o for o in opties if o.afwijking <= 1e-9]
        if onder:
            return max(onder, key=lambda o: (o.dosis, -o.eenheden))
    return min(opties, key=lambda o: (abs(o.afwijking), o.eenheden))
