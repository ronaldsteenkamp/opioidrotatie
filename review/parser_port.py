"""Python-port van de parser in bron/Pharmacom.bas (VertaalCode, Tijden, Hoeveelheid, LeesSterkte, Opioid,
Toedieningsweg, TabelNaam), regel voor regel, zodat de parser ook op Linux/CI te testen is.
Getrouwheid van de port wordt bewaakt door test_referentie.py: elke uitkomst moet gelijk zijn aan wat de
echte VBA gaf (werkmap_uitkomsten.json, gemaakt met excel_harness.ps1).
Bewust NIET verbeterd: fouten in de VBA zitten ook hier in."""
from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Dosering:
    herkend: bool = False
    zo_nodig: bool = False
    per_dag_min: float = 0.0
    per_dag_max: float = -1.0
    max_per_dag: float = -1.0
    aantal: float = 0.0
    aantal_max: float = 0.0
    vorm: str = ""


@dataclass
class Sterkte:
    gevonden: bool = False
    waarde: float = 0.0
    eenheid: str = ""
    per: str = ""
    per_aantal: float = 0.0
    opmerking: str = ""


def _cijfer(c):
    return len(c) == 1 and "0" <= c <= "9"


def _getal(s):
    return s != "" and all(_cijfer(c) for c in s)


def _val(s):
    """VBA Val(): leidend getal, punt als decimaalteken."""
    out = ""
    for c in s:
        if c.isdigit() or (c == "." and "." not in out):
            out += c
        else:
            break
    try:
        return float(out) if out not in ("", ".") else 0.0
    except ValueError:
        return 0.0


def _teken(c):
    return {"K": .25, "H": .5, "Q": .75, "A": 1.5, ")": 10.0, "%": 15.0}.get(c, float(c) if c in "123456789" and c else -1.0)


def hoeveelheid(s):
    if len(s) == 1:
        a = _teken(s)
        return (a, a) if a > 0 else None
    if len(s) == 2:
        a, b = _teken(s[0]), _teken(s[1])
        if a <= 0 or b <= 0 or b <= a:
            return None
        return a, b
    return None


def tijden(k):
    i, wacht, aantal, amin, amax, zo = 0, 0, 0, 0.0, 0.0, False
    while i < len(k):
        if k[i:i + 2] == "MI":
            i += 2
        elif k[i] in "MAN8@(":
            i += 1
        else:
            return None
        wacht += 1
        if i < len(k) and k[i] in "PZ":
            zo = True
            i += 1
        tok = ""
        while i < len(k) and len(tok) < 2 and k[i] in "0123456789KHQ)%":
            tok += k[i]
            i += 1
        if tok:
            h = hoeveelheid(tok)
            if not h:
                return None
            amin += wacht * h[0]
            amax += wacht * h[1]
            aantal += wacht
            wacht = 0
    return (amin, amax, zo) if aantal > 0 and wacht == 0 else None


def vertaal_code(code: str) -> Dosering:
    d = Dosering()
    c = code.strip().replace(" ", "").upper()
    if not c:
        return d
    kern = c
    p = kern.rfind("M") + 1                      # InStrRev, 1-gebaseerd
    if 2 < p < len(kern) and _getal(kern[p:]):
        d.max_per_dag = _val(kern[p:])
        kern = kern[:p - 1]
    vorm = ""
    for v in ("PL", "ZK", "ML", "IJ", "C", "T", "S", "M"):
        if len(kern) > len(v) and kern.endswith(v):
            if not (v == "M" and _teken(kern[-2]) <= 0):
                vorm = v
                kern = kern[:-len(v)]
                break
    d.vorm = vorm
    fmin = fmax = 0.0
    amin = amax = 0.0
    periode = 1.0
    ok = False
    if "-" in kern:
        links, rest = kern.split("-", 1)
        h = hoeveelheid(rest)
        if len(links) == 2 and h:
            amin, amax = h
            fmin = fmax = 1
            periode = 7
            ok = True
    elif len(kern) >= 3 and _cijfer(kern[0]):
        fmin = fmax = _val(kern[0])
        i = 1
        if _cijfer(kern[1]) and kern[2] in ".WPZUSD" and len(kern) >= 4:
            fmax = _val(kern[1])
            i = 2
        sep = kern[i] if i < len(kern) else ""
        if sep and sep in ".WPZUSD":
            h = hoeveelheid(kern[i + 1:])
            if h:
                amin, amax = h
                ok = True
                if sep == "W":
                    periode = 7
                if sep in "PZ":
                    d.zo_nodig = True
    elif kern[:1] in ("P", "Z") and len(kern) >= 2:
        h = hoeveelheid(kern[1:])
        if h:
            amin, amax = h
            ok, d.zo_nodig, fmin, fmax = True, True, 0, -1
    elif kern[:1] == "D" and len(kern) >= 3 and _cijfer(kern[1]):
        h = hoeveelheid(kern[2:])
        if h:
            amin, amax = h
            ok, fmin, fmax, periode = True, 1, 1, _val(kern[1])
    elif kern[:2] == "OD":
        h = hoeveelheid(kern[2:])
        if h:
            amin, amax = h
            ok, fmin, fmax, periode = True, 1, 1, 2
    else:
        t = tijden(kern)
        fmin = fmax = 1
        if t:
            amin, amax, zo = t
            d.zo_nodig = d.zo_nodig or zo
            ok = True
    if not ok:
        return d
    d.herkend = True
    d.aantal, d.aantal_max = amin, amax
    d.per_dag_min = fmin * amin / periode
    if fmax >= 0:
        d.per_dag_max = fmax * amax / periode
    if d.zo_nodig:
        d.per_dag_min = 0
        if d.max_per_dag > 0 and (d.per_dag_max < 0 or d.max_per_dag < d.per_dag_max):
            d.per_dag_max = d.max_per_dag
    return d


# ---------------------------------------------------------------- middel
_ATC = {
    "N02AA01": "Morfine", "N02AA05": "Oxycodon", "N02AA55": "Oxycodon", "N02AA56": "Oxycodon", "N02AA03": "Hydromorfon",
    "N02AB03": "Fentanyl", "N02AE01": "Buprenorfine", "N07BC01": "Buprenorfine", "N07BC51": "Buprenorfine",
    "N02AX02": "Tramadol", "N02AJ13": "Tramadol", "N02AJ14": "Tramadol", "N02AJ15": "Tramadol", "N02AX06": "Tapentadol",
    "N02AC03": "Piritramide", "N07BC02": "Methadon", "N02AC52": "Methadon",
    "R05DA04": "Codeïne", "N02AA59": "Codeïne", "N02AA79": "Codeïne", "N02AJ06": "Codeïne", "N02AJ07": "Codeïne",
    "N02AJ09": "Codeïne",
}
_NAMEN = [("MORF", "Morfine"), ("OXYCOD", "Oxycodon"), ("HYDROMORF", "Hydromorfon"), ("FENTAN", "Fentanyl"),
          ("BUPREN", "Buprenorfine"), ("TRAMAD", "Tramadol"), ("TAPENT", "Tapentadol"), ("PIRITR", "Piritramide"),
          ("METHAD", "Methadon"), ("CODEI", "Codeïne")]
_ZOEK = {v: k for k, v in _NAMEN}


def opioid(atc: str, naam: str) -> str:
    a = atc.strip().upper()
    s = _ATC.get(a[:7], "")
    if s:
        return s
    if a and not a.startswith("N02A") and not a.startswith("N07BC"):
        return ""
    n = naam.upper()
    for zoek, stof in _NAMEN:
        if zoek in n:
            return stof
    return "Onbekend opioïd" if a.startswith("N02A") else ""


def _bevat(t, woorden):
    return any(w in t for w in woorden)


def toedieningsweg(naam: str) -> str:
    n = f" {naam.upper()} "
    if _bevat(n, ("PLEISTER", "PLAKKER", " TTS", "MATRIX")):
        return "pleister"
    if _bevat(n, ("INJ", "INF", "AMPUL", "CASSETTE", " IV ", " SC ")):
        return "parenteraal"
    if _bevat(n, ("NEUS", "NASAL", "INSTANYL", "PECFENT")):
        return "nasaal"
    if _bevat(n, ("SPRAY", "ZUIGT", "SUBL", "BUCC", "FILM", "OROMUC", "SMELT", "LOLLY", "ACTIQ", "ABSTRAL", "EFFENTORA",
                  "BREAKYL")):
        return "transmucosaal"
    if _bevat(n, ("ZETPIL", "SUPP", "RECT")):
        return "rectaal"
    return "oraal"


def tabel_naam(stof: str, weg: str) -> str:
    oraal, par = weg in ("oraal", "rectaal"), weg == "parenteraal"
    if stof in ("Morfine", "Oxycodon", "Hydromorfon"):
        return f"{stof} oraal" if oraal else f"{stof} s.c./i.v." if par else ""
    if stof == "Fentanyl":
        return {"pleister": "Fentanyl pleister", "parenteraal": "Fentanyl s.c./i.v. continu",
                "transmucosaal": "Fentanyl oromucosaal", "nasaal": "Fentanyl nasaal"}.get(weg, "")
    if stof == "Buprenorfine":
        return "Buprenorfine pleister" if weg == "pleister" else ""
    if stof in ("Tramadol", "Tapentadol", "Methadon", "Codeïne"):
        return f"{stof} oraal" if oraal else ""
    if stof == "Piritramide":
        return "Piritramide s.c./i.v." if par else ""
    return ""


def _getal2(s):
    return _val(s.replace(",", "."))


def lees_sterkte(naam: str, stof: str) -> Sterkte:
    s = Sterkte()
    n = naam.replace("µ", "U").replace("μ", "U").upper()
    i = 0
    while i < len(n):
        begin = _cijfer(n[i]) and not (i > 0 and (n[i - 1].isalnum() and n[i - 1].isascii()))
        if begin:
            k = i
            while k < len(n) and n[k] in "0123456789.,/":
                k += 1
            tok, rest = n[i:k], n[k:].lstrip()
            eenh = ""
            for pre, e in (("MICROGRAM", "UG"), ("MCG", "UG"), ("UG", "UG"), ("MG", "MG")):
                if rest.startswith(pre):
                    eenh, rest = e, rest[len(pre):]
                    break
            if eenh:
                getallen, idx, ingr = tok.split("/"), 0, n[:i]
                if "/" in ingr:
                    delen = ingr.split("/")
                    for j, deel in enumerate(delen):
                        if _ZOEK.get(stof, "CODEI") in deel:
                            idx = j
                            break
                    if len(getallen) != len(delen):
                        s.opmerking = "combinatiepreparaat: controleer sterkte"
                elif len(getallen) > 1:
                    s.opmerking = "meerdere sterktes in naam: controleer"
                if idx >= len(getallen):
                    idx = 0
                s.waarde, s.eenheid = _getal2(getallen[idx]), eenh
                s.gevonden = s.waarde > 0
                rest = rest.lstrip()
                if rest.startswith("/"):
                    rest = rest[1:].lstrip()
                    k = 0
                    while k < len(rest) and rest[k] in "0123456789.,":
                        k += 1
                    if k:
                        s.per_aantal = _getal2(rest[:k])
                    rest = rest[k:].lstrip()
                    if rest.startswith("ML"):
                        s.per = "ML"
                    elif rest[:1] in ("U", "H") and s.per_aantal in (0, 1):
                        s.per = "UUR"
                return s
            i = k
        else:
            i += 1
    return s
