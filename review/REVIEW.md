# Review `opioidrotatie` (v1.7, commit `fb11194`)

Review uitgevoerd op 15-10-2026 volgens `codex_review_opioidrotatie.md`. Er zijn geen bestaande bestanden gewijzigd; alles staat onder `review/`.

**Afwijking van de opdracht (in positieve zin):** de opdracht ging uit van Linux zonder Excel. Deze review is gedaan op Windows **mét Excel**. Alle reken- en parsercasussen zijn daarom door de **echte, gebouwde werkmap** gehaald (Excel via COM, alleen-lezen geopend en nooit opgeslagen), inclusief de echte VBA. Eerst is gecontroleerd dat de `.xlsm` in de repo-map formule-voor-formule gelijk is aan een verse build uit `bron/werkmap_opbouw.py`: 0 verschillen.

Zekerheid: **bevestigd** = uitgevoerd en aangetoond, **waarschijnlijk** = afgeleid uit de code, **vraag** = moet de auteur beantwoorden.

---

## Status na herstel (15-10-2026, na de review)

K-01, K-02 en K-03 zijn opgelost in `bron/`, de werkmap is opnieuw gebouwd en de casussen zijn opnieuw door Excel gehaald. pytest: **245 geslaagd, 11 xfail** (de overige bevindingen). Bewijs per bevinding: de bijbehorende tests zijn van `xfail` naar gewone tests gegaan.

| Bevinding | Wijziging | Effect (casus) |
|---|---|---|
| K-01 | `werkmap_opbouw.py`, Omzetten doorbraak: `ROUNDDOWN(ROUND(L/C,9),0)` in plaats van `ROUND` | Z11: in plaats van 10 mg (+100%) nu "lager dan de kleinste sterkte". Z12: 10 mg (−33%) in plaats van 20 mg (+33%). Z02: geen advies meer in plaats van +33%. |
| K-02 | `werkmap_opbouw.py`, pompstand en pompstand bij de andere concentratie: naar beneden afronden (`ROUNDDOWN(ROUND(x/stap,9),0)*stap`) | P17: 0,10 ml/u (−20%) in plaats van 0,15 (+20%). P11/P13: geen 0,01 of 0,10 (+44%) meer, maar "te laag". De afwijking is nooit meer positief. |
| M-02 (meegenomen) | De hoofdtegel toont "te laag" (met verwijzing naar stap 5) in plaats van "0,00 ml/uur" | Hoort bij K-02: door het naar beneden afronden komt pompstand 0 vaker voor. |
| K-03 | `Pharmacom.bas`, `EigenCode` leest `Doseercodes!B9:B58` (waarden in C–F) | `EIGEN1` en het voorbeeld `VOORB1` worden herkend "(eigen tabel)". De layouttest is groen. |

Docs aangepast: README (afrondingsrichting), Toelichting (rekenmethode + voorbeeld 0,229 → 0,22 ml/u), de notitie bij de doorbraak op Omzetten. Uitgebracht als **v1.7.1**, met nieuwe `docs/*.png`.

Let op na K-02: bij lage doses verschijnt nu vaker "te laag" in plaats van een (te hoge) pompstand. Het voorstel voor de andere concentratie geldt alleen als die binnen 10% blijft. Bij 5 mg morfine oraal geeft 1 mg/ml 0,06 ml/u (−14%), dus de melding is dan algemeen.

---

## Samenvatting

De rekenkern is degelijk. Op alle 33 numerieke pompcasussen geeft de werkmap **exact** (verschil 0) hetzelfde als de onafhankelijke Python-referentie. Alle factoren in de README komen overeen met de werkmap, en de waarschuwingen gaan bij de juiste grenzen af. De parser is verrassend robuust: 37 doseercodes en 30 etiketnamen, en de Python-port gedraagt zich op elke testcasus precies als de VBA. De risico's zitten in de randen: afronding, het preparaatadvies en een paar plekken waar de software stil iets anders doet dan de gebruiker denkt.

**Top-5 risico's**
1. **K-01** – *Omzetten* adviseert een doorbraakdosis tot **+100%** boven 1/6 van de dagdosis (bijv. 10 mg morfine waar 5 mg berekend is) en geeft daarbij **geen** waarschuwing.
2. **K-02** – De pompstand wordt afgerond op de *dichtstbijzijnde* stap, dus ook **omhoog**: +44% bij lage doses en +20% bij stap 0,05. De tegel noemt dat "instellen op de pomp". Er staat wel een waarschuwing bij.
3. **K-03** – Eigen doseercodes op tabblad *Doseercodes* worden **nooit gelezen**. De VBA leest `A7:E56`, maar het tabblad zet de codes in `B9:F58`. Een correctie door de gebruiker wordt dus stil genegeerd.
4. **H-03 / H-04** – Een eGFR die als tekst getypt is (`<15`) geeft geen nierwaarschuwing. Hydromorfon zonder ATC-code wordt als **morfine** ingevuld (5× te lage OME) en wordt niet oranje.
5. **H-05 (vraag)** – De VBA gebruikt de kolommen *Actief* en *Startdatum* niet. Als de medicatiestatus niet-actieve of dubbele regels toont, telt de OME die mee en wordt de pompdosis te hoog.

---

## Bevindingen

Ernst: **Kritiek** = kan tot een verkeerde dosis leiden · **Hoog** · **Middel** · **Laag**.

| ID | Ernst | Categorie | Bestand:regel | Beschrijving | Zekerheid | Aanbeveling |
|---|---|---|---|---|---|---|
| K-01 | Kritiek | Klinisch | `bron/werkmap_opbouw.py:1185-1189`, `:1212-1214` | Het doorbraakadvies op *Omzetten* rondt het aantal eenheden af naar het dichtstbijzijnde getal (`ROUND`) en kiest de kleinste afwijking, ook als die erboven ligt. Daarbij geldt geen bovengrens en is er geen waarschuwing: de >10%-controle (r. 1212) kijkt alleen naar het onderhoudsadvies. Voorbeeld Z11: 40 OME, bijwerkingen → 30 mg → doorbraak 1/6 = 5 mg, maar het advies is **morfine 10 mg P1TM6 (+100%)**, dus tot 60 mg/dag extra bij een onderhoud van 30 OME. Z02 en Z12: +33%. | bevestigd | Kies het doorbraakpreparaat nooit boven de berekende dosis (afronden naar beneden). Valt het onder de kleinste sterkte, toon dan "lager dan kleinste sterkte" in plaats van een advies. Pas de >10%-melding ook toe op de doorbraak. |
| K-02 | Kritiek | Klinisch | `bron/werkmap_opbouw.py:495`, `:509-510` | `ROUND(exact/stap,0)*stap` rondt half-omhoog af. P11: 5 mg morfine oraal, 10 mg/ml, stap 0,01 → exact 0,0069 ml/u → **0,01 ml/u = 2,4 mg/24 u in plaats van 1,67 (+44%)**. P13 (1 mg/ml, stap 0,1): +44%. P17 (stap 0,05): 0,125 → **0,15 ml/u (+20%)**. De waarschuwing "Afronding wijkt X% af" gaat wel af (>10%), en bij P11 stelt de werkmap terecht 1 mg/ml voor. Maar de hoofdtegel toont de verhoogde stand als de in te stellen waarde. | bevestigd | Rond af naar beneden (`ROUNDDOWN`/`FLOOR`), of toon bij >10% afwijking in de tegel geen pompstand maar "kies andere concentratie". README en Toelichting moeten de afrondingsrichting noemen. |
| K-03 | Kritiek | Parser | `bron/Pharmacom.bas:680-707` ↔ `bron/werkmap_opbouw.py:748-757` | `EigenCode` zoekt de code in `Doseercodes!A7:A56`, met de waarden in B–E. Het tabblad heeft de kop in rij 8 en de codes in **kolom B, rij 9-58** (waarden in C–F). Aangetoond: `EIGEN1` in B10 geeft "niet herkend", het voorbeeld `VOORB1` ook. Alleen een code in kolom A wordt gelezen, en dan met de verkeerde kolommen. Een gebruiker die een stil verkeerd vertaalde code (zie H-06) wil corrigeren, krijgt nog steeds de verkeerde vertaling, zonder oranje. README r. 89 belooft deze functie. | bevestigd | Lees `B9:B58` met `Offset` 1..4 (of beter: een benoemd bereik `tblDoseercodes` dat Python aanmaakt en VBA gebruikt). Neem de layouttest uit `review/test_referentie.py` op (`test_layout_doseercodes`). |
| H-01 | Hoog | Klinisch | `bron/werkmap_opbouw.py:1115-1123` | Het onderhoudsadvies in *Omzetten* kiest de kleinste absolute afwijking, ook boven de berekende dosis. Z02: 18,75 µg/u → **2× fentanyl 12 = 24 µg/u (+28%)**. Z13: 8,3 µg/u → 12 µg/u (+44%). Er komt wel de melding "wijkt meer dan 10% af", maar het groene advies staat er gewoon. De opdracht vroeg: "nooit >10% boven de berekende dosis". Dat wordt **niet** gehaald. | bevestigd | Bij een afwijking van meer dan +10% geen groen advies. Bied de beste optie ≤ berekend aan en toon de hogere optie alleen ter vergelijking. Overweeg een asymmetrische score (overdosering zwaarder laten wegen). |
| H-02 | Hoog | Klinisch | `bron/werkmap_opbouw.py:1125-1129`, `:1211` | Ligt de nieuwe dosis boven de max. dagdosis, dan toont de werkmap het preparaat en de doseercode toch. Z09: 300 OME → tramadol → **"TRAMADOL 200MG 2.3T + 150MG 2.1T = 1500 mg/24 u"** (max. 400), met alleen een melding eronder. | bevestigd | Bij `DAG > MAXT` geen preparaat en code tonen, maar "boven max. dagdosis – kies een ander middel". |
| H-03 | Hoog | Klinisch | `bron/werkmap_opbouw.py:453-454`, `:610` | De eGFR-cel D22 heeft geen gegevensvalidatie. De waarschuwing gebruikt `ISNUMBER(D22)`. Typt de gebruiker `<15` (zoals het lab rapporteert), dan komt er **geen** waarschuwing (P24). Bij uitlezen uit Pharmacom gaat het wel goed: de VBA zet `<15` om naar 15 en maakt de cel oranje (V06). | bevestigd | Validatie op D22 (decimaal ≥ 0), of de waarschuwing laten afgaan op `VALUE(SUBSTITUTE(SUBSTITUTE(D22,"<",""),">",""))`. |
| H-04 | Hoog | Parser | `bron/Pharmacom.bas:344-353` | Zonder ATC-code (of met een onbekende N02A-code) wordt het opioïd op naam herkend. `"MORF"` staat in de lijst vóór `"HYDROMORF"`, dus *HYDROMORFON CAPSULE MGA 4MG* wordt **Morfine oraal**. V02: 2.1C → morfine 8 mg = 8 OME in plaats van 40 OME (5× te laag), **niet oranje**. | bevestigd | `HYDROMORF` vóór `MORF` zetten (langste naam eerst). Herkenning op naam altijd oranje maken. |
| H-05 | Hoog | Koppeling | `bron/Pharmacom.bas:149-201`; `Pharmacom_medicatiestatus.ps1:83,118` | Het script leest *Actief* en *Startdatum* wel uit, maar `Verwerk` gebruikt ze niet. Alleen *Th. einddatum < vandaag* sluit een regel uit. Als de medicatiestatus ook niet-actieve, toekomstige of dubbele regels (herhaalrecepten) toont, telt de OME die allemaal mee → te hoge pompdosis. | vraag | Vraag: welke regels toont de medicatiestatus precies, en wat staat er in *Actief*? Sluit niet-actieve regels uit, en meld dubbele regels van hetzelfde middel oranje. |
| H-06 | Hoog | Parser | `bron/Pharmacom.bas:542-551` | De scheidingstekens `U`, `S` en `D` na de frequentie worden behandeld als "keer per dag". `4U1T` → **4 tabletten/24 u**, niet oranje. Betekent `4U` "elke 4 uur" (6×/dag), dan is dat stil 33% te laag. Het omgekeerde kan ook bij andere codes. | vraag | Controleer elke scheider tegen de Centrale doseertabel. Wat niet zeker is, maak oranje ("nakijken"). |
| M-01 | Middel | Klinisch | `bron/werkmap_opbouw.py:381-382`, `:398` | Een negatieve dosis (geplakt, wat de validatie omzeilt) verlaagt de totale OME zonder melding. P26: 120 morfine + (−40 oxycodon) → **60 OME**, Controle "✓ OK". | bevestigd | Controlekolom H: `D<0` → "Negatieve dosis". Of de totaal-OME laten weigeren bij negatieve regels. |
| M-02 | Middel | Klinisch | `bron/werkmap_opbouw.py:520` | Rondt de pompstand af naar 0, dan toont de hoofdtegel **"0,00 ml/uur"** en bolus 0,0 mg (P12, P14). De waarschuwing (−100%) staat er wel. | bevestigd | Bij `D47=0` in de tegel "te lage dosis voor deze stap/concentratie" tonen. |
| M-03 | Middel | Parser | `bron/Pharmacom.bas:380-384` | *FENTANYL SPRAY NASAAL* bevat niet de tekst `NASAL` (wel `NASAAL`) en valt op `SPRAY` → **Fentanyl oromucosaal** (factor 0,05 in plaats van 0,3, dus 6× te laag). Wel oranje, maar het middel staat er verkeerd. | bevestigd | `NASAAL` toevoegen en nasaal vóór `SPRAY` controleren. |
| M-04 | Middel | Parser | `bron/Pharmacom.bas:344`, `:376` | (a) Merknaam zonder ATC (*PALLADON SLOW 8MG*) → "geen opioïd": **stil weggelaten**, niet in de samenvatting (V11). (b) *TRANSTEC 35 MICROGRAM/UUR* (zonder het woord PLEISTER) → weg "oraal" → niet ingevuld; dit wordt wel gemeld. | bevestigd | Regels met een lege ATC-code altijd in de samenvatting noemen ("niet beoordeeld"). Merknamen (Transtec, Durogesic, Palladon, OxyContin, …) aan de herkenning toevoegen. |
| M-05 | Middel | Koppeling | `bron/Pharmacom.bas:101-106` | De eGFR-tekst wordt teken voor teken op cijfers gefilterd. `45 (CKD-EPI 2009)` wordt **452009** (V07): oranje, maar de nierwaarschuwing gaat niet af. | bevestigd | Alleen het eerste getal nemen (regex `[<>]?\s*(\d+([.,]\d+)?)`), en onwaarschijnlijke waarden (>200) weigeren. |
| M-06 | Middel | Koppeling | `Pharmacom_medicatiestatus.ps1:158-174` | De eGFR-regel wordt gevonden in de tabel *Naam* en via hetzelfde **rijnummer** gelezen uit een andere tabel (uitslagen). Wijkt de rijvolgorde af (sorteren, gevirtualiseerde rijen, verborgen regels), dan wordt stil een andere uitslag gelezen. | waarschijnlijk | Controleer het aantal rijen van beide tabellen, en lees de eenheid of normaalwaarde mee als controle. |
| M-07 | Middel | Privacy | `bron/Pharmacom.bas:58-63`, `:169`, `:114` | De werkmap bevat na het uitlezen patiëntgegevens: venstertitel in B4 (vermoedelijk met naam), medicatie in kolom I en de eGFR als opmerking. Wissen gebeurt alleen met de knop. Er is geen `Workbook_BeforeSave`/`BeforeClose`, en AutoHerstel kan een kopie bewaren in `%APPDATA%\Microsoft\Excel`. | waarschijnlijk | In `ThisWorkbook`: bij `BeforeClose` blok 1, B4 en eGFR wissen. `BeforeSave` waarschuwen of annuleren als B4 ≠ plaatshouder. `Me.EnableAutoRecover = False` bij openen. |
| M-08 | Middel | Privacy | `Pharmacom_naar_Excel.ps1:72-82` | Het tijdelijke bestand in `%TEMP%` wordt verwijderd na `ReadAllText`, maar niet in `try/finally`. Bij een fout of afbreken tussen schrijven en lezen blijft medicatie mét venstertitel staan. README zegt "direct verwijderd". | waarschijnlijk | `try { … } finally { Remove-Item -LiteralPath $Tmp, "$Tmp.tmp" -EA 0 }`. Ruim bij de start oude `opioidrotatie_*.txt` op. Beter nog: stdout gebruiken in plaats van een bestand. |
| M-09 | Middel | Test | `bron/Test_werkmap.ps1:10-77` | De test **print** alleen en controleert niets. K-03 (`VOORB1` wordt "niet herkend") staat al in de uitvoer van de bestaande test, maar valt niemand op. | bevestigd | Asserties toevoegen (verwacht ↔ werkelijk, exitcode ≠ 0 bij fouten), of `review/test_referentie.py` + `excel_harness.ps1` gebruiken. |
| M-10 | Middel | Regelgeving | `README.md:9-11` | Zie het hoofdstuk Regelgeving. "Geen medisch hulpmiddel" is een claim van de maker. Voor de MDR telt het beoogde gebruik: patiëntspecifieke doseringen berekenen. | te laten beoordelen | Laten beoordelen (zie onder). |
| L-01 | Laag | Docs/Klinisch | `bron/werkmap_opbouw.py:201-202` | Buprenorfine: factor 2,4 geeft voor 52,5 µg/u **126** mg OME. De opmerking in dezelfde rij zegt "KNMP-tabel: 52,5 µg/uur = 120 mg". Intern inconsistent (5%). | bevestigd | Controleer tegen de KNMP-tekst. De factor of de opmerking aanpassen. |
| L-02 | Laag | Klinisch | `bron/werkmap_opbouw.py:619-620` | De melding "spuit binnen 24 uur leeg" stelt `MAX(conc1,conc2)` voor, ook als dat de huidige concentratie is. P15: "… bij 10,0 mg/ml. **Overweeg 10 mg/ml** of een groter volume." | bevestigd | Alleen een andere concentratie voorstellen als die hoger is dan de huidige. |
| L-03 | Laag | Klinisch | `bron/werkmap_opbouw.py:614` | "Volledige equianalgetische dosis …" verschijnt ook bij een lege stap 1 (P25). | bevestigd | `AND(G16>0, …)` toevoegen. |
| L-04 | Laag | Klinisch | `bron/werkmap_opbouw.py:400-406` | Tekst in het dosisveld (`60 mg`) geeft terecht `#WAARDE!` in de uitkomst, maar de Controle-kolom zegt "✓ OK" (P27). | bevestigd | `NOT(ISNUMBER(D))` → "Geen getal". |
| L-05 | Laag | Docs | `README.md:89,99-110`; `bron/Pharmacom.bas:15` | De README-tabel mist piritramide (factor 1,5) en codeïne (0,1), die wel in de werkmap staan. README r. 89 belooft eigen codes (K-03). De VBA-kop zegt "er worden geen toetsen naar Pharmacom gestuurd", terwijl `Pharmacom_naar_Excel.ps1:104,115-118` F3 en Alt+I/Alt+D stuurt (de README zegt dat wel correct). | bevestigd | Docs gelijktrekken. |
| L-06 | Laag | Kwaliteit | `bron/werkmap_opbouw.py:139-143,163`; `bron/Pharmacom.bas:20-22,687,745-754` | `S_BOLLO` = 1/6 (de *hoge* variant) en `S_BOLHI` = 10%: de namen zijn omgekeerd. Celadressen staan dubbel in Python en VBA, met alleen een commentaar als koppeling (de oorzaak van K-03). `LeesUtf8` is dode code. | bevestigd | Benoemde bereiken (`Name`) laten aanmaken door Python en gebruiken in VBA. Namen omdraaien. Dode code weg. |
| L-07 | Laag | Koppeling | `Pharmacom_naar_Excel.ps1:45-55`, `:98` | Tussen de controle "Pharmacom is voorgrond" en `SendWait` zit een klein tijdvenster (TOCTOU). Verder matcht `'FOUT*niet gevonden*'` ook "Java Access Bridge niet gevonden" (`Pharmacom_medicatiestatus.ps1:37`), waardoor onnodig 2× F3 wordt gestuurd. | waarschijnlijk | Op een specifieke foutcode matchen (bijv. `FOUT<tab>GEEN_F3`). |
| L-08 | Laag | Parser | `bron/Pharmacom.bas:506-512` | Een tijdcode zonder vormletter met `M` (ochtend) op het eind, zoals `A1M1`, wordt gelezen als "niet meer dan 1" → **1/24 u in plaats van 2**, niet oranje. Komt zelden voor, omdat codes normaal op een vorm eindigen en chronologisch zijn (`M1A1` gaat goed). | bevestigd (port, die op alle 37 codes gelijk is aan de VBA) | `M`+getal alleen als maximum lezen als er een vorm vóór staat, of als de rest geen tijdcode is. |
| L-09 | Laag | Security | `Installeer.ps1:70`; `Pharmacom_naar_Excel.ps1:136` | `Add-Content` voegt `assistive_technologies=` toe zonder te kijken of er al een andere waarde staat. Excel geopend via COM draait macro's ongeacht de vertrouwensinstelling (AutomationSecurity standaard laag). Dat is beperkt, omdat het pad vastligt naast het script. | waarschijnlijk | Bestaande regel vervangen in plaats van toevoegen. Optioneel `$Xl.AutomationSecurity` expliciet zetten en de werkmap alleen uit `$PSScriptRoot` openen (gebeurt al). |
| L-10 | Laag | Kwaliteit | `bron/werkmap_opbouw.py:172-216`, `:886` | De bron staat per factor vermeld, maar zonder raadpleegdatum per factor (alleen globaal in de Toelichting, 14-10-2026). Pleistersterkte "37" en "12" µg/u: de handelsproducten zijn 37,5 en 12,5 µg/u. | vraag | Kolom "geraadpleegd op" per factor. Sterktes controleren tegen de G-Standaard. |
| L-11 | Laag | Koppeling | `bron/Pharmacom.bas:766-770` | Een tweecijferig jaartal (`31-12-30`) gaat via `DateSerial` en de Windows-instelling voor jaartallen (hier: 2030). Op een pc met een andere instelling wordt dat 1930: de regel geldt dan als "beëindigd" en wordt overgeslagen (wel gemeld). | waarschijnlijk | Bij een jaartal van 2 cijfers zelf 20xx gebruiken, of oranje melden. |

---

## Rekencontrole (stap 2)

*Uitgevoerd*: elke casus is in de echte werkmap gezet (Excel via COM; `review/excel_harness.ps1`) en vergeleken met `review/referentie.py`. Die is geschreven vanuit de README en neemt de formules **niet** over. *Verschil* is het grootste absolute verschil over OME, dagdosis, ml/uur, werkelijke dagdosis en afwijking. Standaard: 10 mg/ml, stap 0,01 ml/u, spuit 50 ml, onvoldoende pijnstilling (100%).

| ID | Casus | Referentie: OME → s.c. mg/24 u → ml/uur (stap) → werkelijk mg/24 u, afw. | Werkmap (tegels) | Verschil | Oordeel |
|---|---|---|---|---|---|
| P01 | 60 mg morfine oraal, onvoldoende effect | 60 → 20,0 → 0,08 (0,01) → 19,2, −4% | 0,08 ml/u · 19,2 mg | 0 | klopt (bolus 3,2 mg = 1/6) |
| P02 | fentanylpleister 25 µg/u, bijwerkingen (75%) | 60 → 15,0 → 0,06 → 14,4, −4% | 0,06 · 14,4 | 0 | klopt; pleisterwaarschuwing aanwezig |
| P03 | oxycodon 2×20 + P1CM6 5 mg (max 30) | 105 → 35,0 → 0,15 → 36,0, +3% | 0,15 · 36,0 | 0 | klopt |
| P04 | morfine 60 + tramadol 200 + fentanyl 12 | 128,8 → 42,9 → 0,18 → 43,2, +1% | 0,18 · 43,2 | 0 | klopt |
| P05 | methadon 15 mg | 60 → 20,0 → 0,08 → 19,2 | 0,08 · 19,2 | 0 | klopt (factor 4) |
| P06 | methadon 20 mg | 80 → 26,7 → 0,11 → 26,4 | 0,11 · 26,4 | 0 | klopt (factor 4) |
| P07 | methadon 21 mg | 126 → 42,0 → 0,18 → 43,2 | 0,18 · 43,2 | 0 | klopt (factor 6; sprong +58% t.o.v. 20 mg, gedocumenteerd) |
| P08 | methadon 50 mg | 300 → 100 → 0,42 → 100,8 | 0,42 · 100,8 | 0 | klopt (factor 6) |
| P09 | methadon 51 mg | 408 → 136 → 0,57 → 136,8 | 0,57 · 136,8 | 0 | klopt (factor 8) |
| P10 | methadon 20,5 mg | 123 → 41,0 → 0,17 → 40,8 | 0,17 · 40,8 | 0 | klopt (>20 → factor 6) |
| P11 | zeer laag: morfine 5 mg, stap 0,01 | 5 → 1,67 → **0,01** → 2,4, **+44%** | 0,01 · 2,4 | 0 | **K-02**: rondt omhoog af; waarschuwing + "kies 1 mg/ml: 0,07" |
| P12 | zeer laag: morfine 5 mg, stap 0,1 | 5 → 1,67 → 0,00 → 0, −100% | 0,00 · 0,0 | 0 | **M-02**: 0,00 in de tegel; waarschuwing aanwezig |
| P13 | zeer laag: morfine 5 mg, 1 mg/ml, stap 0,1 | 5 → 1,67 → **0,10** → 2,4, **+44%** | 0,10 · 2,4 | 0 | **K-02** |
| P14 | zeer laag: morfine 2 mg | 2 → 0,67 → 0,00 → 0, −100% | 0,00 · 0,0 | 0 | **M-02** |
| P15 | zeer hoog: oxycodon 800 (1200 OME), spuit 20 | 1200 → 400 → 1,67 → 400,8 | 1,67 · 400,8 | 0 | klopt; hoge dosis + spuit-melding (**L-02**: tekst) |
| P16 | morfine 90, stap 0,1 | 90 → 30 → 0,10 → 24,0, −20% | 0,10 · 24,0 | 0 | klopt; stelt 1 mg/ml 1,30 ml/u voor |
| P17 | morfine 90, stap 0,05 | 90 → 30 → **0,15** → 36,0, **+20%** | 0,15 · 36,0 | 0 | **K-02**: 0,125 precies op de helft → omhoog |
| P18 | morfine 90, 1 mg/ml, stap 0,05, spuit 20 | 90 → 30 → 1,25 → 30,0, 0% | 1,25 · 30,0 | 0 | klopt; spuit leeg na 16 u: melding |
| P19 | morfine 90, 1 mg/ml, stap 0,1 | 90 → 30 → 1,30 → 31,2, +4% | 1,30 · 31,2 | 0 | klopt |
| P20 | eGFR 29 | – | – | 0 | waarschuwing eGFR < 30: ok |
| P21 | eGFR 30 | – | – | 0 | geen waarschuwing (grens `<30`): conform README |
| P22 | eGFR ontbrekend | – | – | 0 | geen waarschuwing: ok |
| P23 | eGFR `>90` (tekst) | – | – | 0 | geen waarschuwing: correct |
| P24 | eGFR `<15` (tekst, getypt) | – | – | 0 | **H-03**: geen waarschuwing |
| P25 | lege invoer | 0 | tegels leeg | 0 | ok; **L-03** |
| P26 | negatieve dosis | (120 − 60 =) 60 | 0,08 · 19,2 | 0 | **M-01**: stil verlaagd |
| P27 | tekst `60 mg` in het dosisveld | n.v.t. | #WAARDE! | – | veilig zichtbaar; **L-04** |
| P28 | `"60"` als tekst (geplakt) | 60 | 0,08 · 19,2 | – | Excel rekent het als getal: ok |
| P29 | tramadol 500 (max 400) | 100 → 33,3 → 0,14 → 33,6 | 0,14 · 33,6 | 0 | melding boven max: ok |
| P30 | precies 240 OME | 240 → 80 → 0,33 → 79,2 | 0,33 · 79,2 | 0 | "hoge dosis" bij ≥240: ok |
| P31 | bijwerkingen + 25% extra reductie | 120 → 22,5 → 0,09 → 21,6 | 0,09 · 21,6 | 0 | klopt (0,75 × 0,75) |
| P32 | "Anders" zonder % | 60 → 0 → 0,00 | 0,00 · 0,0 | 0 | melding "vul % in": ok |
| P33 | fentanyl s.c./i.v. 25 µg/u | 180 → 60 → 0,25 → 60,0 | 0,25 · 60,0 | 0 | klopt (factor 7,2) |
| P34 | fentanyl oromucosaal 1200 µg/24 u | 60 → 20 → 0,08 → 19,2 | 0,08 · 19,2 | 0 | klopt; melding "indicatief" |
| P35 | buprenorfine 52,5 µg/u | 126 → 42 → 0,18 → 43,2 | 0,18 · 43,2 | 0 | klopt met factor 2,4; **L-01** |

Decimale komma of punt: invoer via Excel volgt de landinstelling; via de VBA wordt `Getal2` gebruikt (komma → punt), en dat is getest met `37,5`, `52,5` en `2,6`. Deling door nul is overal afgevangen (`N(D26)>0`, `D47>0`, `D44>0`). Eenheden µg/u ↔ mg/24 u (×24/1000) en ml ↔ mg kloppen in alle casussen.

**Omzetten** (factor, staffel en preparaatkeuze):

| ID | Casus | Referentie: nieuwe dosis | Werkmap | Werkmap-advies (groen) | Doorbraakadvies | Oordeel |
|---|---|---|---|---|---|---|
| Z01 | 165 OME → fentanylpleister | 68,75 µg/u | 68,75 | FENTANYL 75: D31PL = 75 (+9%) | morfine 20 mg P1TM6 (−27%) | ok |
| Z02 | 60 OME, bijwerkingen (45) → fentanylpleister | 18,75 | 18,75 | **2× fentanyl 12 = 24 (+28%)** | morfine 10 mg (**+33%**) | **H-01**, **K-01** |
| Z03 | 60 OME → oxycodon oraal | 40 mg | 40 | oxycodon MGA 20 mg 2.1T (0%) | oxycodon 5 mg P1CM6 (−25%) | ok |
| Z04 | 105 OME → hydromorfon oraal | 21 mg | 21 | 4 mg 2.3C = 24 (+14%) | 2,6 mg (−26%) | melding >10%: ok |
| Z05 | 60 OME → methadon | 15 mg | 15 | 2 mg 2.4T = 16 (+7%) | geen preparaat | ok |
| Z06 | 100 OME → methadon (gat in de staffel) | 20 mg | 20 | 10 mg 2.1T = 20 | – | ok (lagere trede, conform tekst) |
| Z07 | 350 OME → methadon (gat in de staffel) | 50 mg | 50 | 20 mg 2.1T + 5 mg 2.1T = 50 | – | ok |
| Z08 | 500 OME → methadon | 62,5 mg | 62,5 | 10 mg 2.3T = 60 (−4%) | – | ok |
| Z09 | 300 OME → tramadol | 1500 mg | 1500 | **200 mg 2.3T + 150 mg 2.1T = 1500** | – | **H-02** (max 400) |
| Z10 | morfine 45 → oxycodon 3×/dag | 30 mg | 30 | 10 mg 3.1T (0%) | 5 mg (0%) | ok |
| Z11 | 40 OME, bijwerkingen (30) → fentanylpleister | 12,5 | 12,5 | fentanyl 12 (−4%) | **morfine 10 mg = +100%** | **K-01** |
| Z12 | 90 OME → buprenorfinepleister | 37,5 | 37,5 | 35 µg/u (−7%) | morfine 20 mg (**+33%**) | **K-01** |
| Z13 | 20 OME → fentanylpleister | 8,33 | 8,33 | **fentanyl 12 (+44%)** | lager dan kleinste sterkte | **H-01** (wel melding) |
| Z14 | 1300 OME → fentanylpleister | 541,7 | 541,7 | geen advies (>4 eenheden) | – | veilig |
| – | Tabblad *Preparaten* leeg | – | – | "Geen preparaten voor dit middel op tabblad Preparaten." | "Geen doorbraakpreparaat …" | ok, nette melding |

De methadon-omrekening terug is onafhankelijk getest: over 1–999 OME geeft de omgekeerde staffel nooit een methadondosis die via de staffel méér OME oplevert (`test_methadon_omgekeerd_nooit_boven_ome`).

---

## Parsercontrole (stap 3)

*Uitgevoerd*: elke code via de echte VBA-functie `VertaalCodeTekst` en elke naam via `SterkteTekst`, plus 11 end-to-end-scenario's via `Verwerk` (verzonnen regels). De Python-port (`review/parser_port.py`) geeft op **alle** 37 codes en 30 namen hetzelfde als de VBA. "Bedoeld" is mijn lezing van de Centrale doseertabel-systematiek zoals die in de code en Toelichting beschreven staat. Die is **niet** geverifieerd tegen de officiële tabel.

**Doseercodes**

| Code | Bedoeld (eenheden/24 u) | Parser (VBA) | Oordeel |
|---|---|---|---|
| `2.1T` | 2 | 2× per dag 1 tablet = 2 | goed |
| `3.1T` | 3 | 3 | goed |
| `6.1T` | 6 | 6 | goed |
| `2.2T` | 4 | 4 | goed |
| `3.12T` | 3–6 | 3–6, laagste ingevuld | goed, oranje |
| `2.H1T` | 1–2 | 1–2 | goed, oranje |
| `2.1C` / `2.1S` | 2 | 2 (capsule / zetpil) | goed |
| `P1CM6` | zo nodig, max 6 | max 6 → doorbraak | goed, oranje + waarschuwing |
| `P2TM6` | zo nodig, max 6 | max 6 | goed |
| `4P1TM4` | zo nodig, max 4 | max 4 | goed |
| `P1C` / `Z1T` | zo nodig, geen max | dosis leeg | goed, oranje |
| `3Z1T` | zo nodig, max 3 | max 3 | goed |
| `P5MLM30` | zo nodig, max 30 ml | max 30 ml | goed |
| `2.1TM2` | 2 | 2 | goed |
| `MA1T` | 2 | 2 | goed |
| `M1A2T` | 3 | 3 | goed |
| `MAN1T` | 3 | 3 | goed |
| `A1M1T` | 2 | 2 | goed |
| `A1M1` (zonder vorm) | 2 | **1** ("niet meer dan 1") | **stil verkeerd** (L-08, port) |
| `D31PL` | 1 pleister per 3 dagen | 1 → dosis = sterkte | goed |
| `D32PL` | 2 pleisters | 2 → 2× sterkte | goed, oranje bij bereik |
| `D41PL` / `1W1PL` | per 4 dagen / week | goed | goed |
| `1.1PL` | 1 | 1 | goed |
| `OD1T` | 0,5 | 0,5 | goed |
| `ZO-1T` | 1/7 | 1/7 | goed |
| `4.5ML` (drank 10 mg/5 ml) | 20 ml = 40 mg | 40 mg | goed (end-to-end) |
| `2D1T` | 2 | 2 | goed |
| `3D1-2T` | onbekend | niet herkend | goed (oranje) |
| `1D1PL72U` | onbekend | niet herkend | goed (oranje, 1 pleister aangenomen) |
| `3U1T` / `4U1T` | **onzeker** | 3 / 4 per dag | **mogelijk stil verkeerd** (H-06) |
| `12.1T` | 1–2 | 1–2 | goed |
| `10.1T` | onbekend | 1 (max 0) | twijfel; niet oranje |
| `Z1M4` / `Z12M8` | zo nodig, max 4 / 8 | max 4 / 8 | goed |
| `VOORB1` / eigen code in *Doseercodes* | uit eigen tabel | **niet herkend** | **K-03** |

**Etiketnamen**

| Etiketnaam | ATC | Bedoeld | Parser (VBA) | Oordeel |
|---|---|---|---|---|
| OXYCODON TABLET MGA 20MG | N02AA05 | Oxycodon oraal 20 | idem | goed |
| OXYCODON/NALOXON TABLET MGA 10/5MG | N02AA55 | Oxycodon oraal 10 | 10 | goed (positie van het opioïd) |
| TARGINACT TABLET MGA 20/10MG | N02AA55 | Oxycodon oraal 20 | 20, "meerdere sterktes" | goed, oranje |
| MORFINE DRANK 10MG/5ML | N02AA01 | 2 mg/ml | 10 MG/5ML → 2 mg/ml | goed |
| MORFINE HCL DRANK 2MG/ML | N02AA01 | 2 mg/ml | idem | goed |
| MORFINE INJVLST 10MG/ML | N02AA01 | Morfine s.c./i.v. | idem | goed |
| MORFINE ZETPIL 10MG | N02AA01 | Morfine oraal (1:1) | Morfine oraal | goed, maar rectaal = oraal is niet gedocumenteerd (vraag) |
| FENTANYL PLEISTER 12UG/UUR | N02AB03 | 12 | 12 | goed |
| FENTANYL PLEISTER 37,5UG/UUR | N02AB03 | 37,5 | 37,5 | goed |
| FENTANYL MATRIXPLEISTER 25MICROGRAM/UUR | N02AB03 | 25 | 25 | goed |
| DUROGESIC 50 PLEISTER 50UG/UUR | N02AB03 | 50 | 50 | goed |
| FENTANYL 12 MICROGRAM/UUR PLEISTER | N02AB03 | 12 | 12 | goed |
| BUPRENORFINE PLEISTER 52,5UG/UUR 3 DAGEN | N02AE01 | 52,5 | 52,5 | goed |
| TRANSTEC 35 MICROGRAM/UUR | N02AE01 | Buprenorfine pleister 35 | weg "oraal" → niet ingevuld | gemeld (M-04b) |
| BUPRENORFINE TABLET SUBLINGUAAL 0,4MG | N07BC01 | niet automatisch | niet ingevuld | goed |
| ABSTRAL TABLET SUBLINGUAAL 100MICROGRAM | N02AB03 | Fentanyl oromucosaal | idem | goed, oranje |
| INSTANYL NEUSSPRAY 50UG/DOSIS | N02AB03 | Fentanyl nasaal | idem | goed, oranje |
| FENTANYL SPRAY NASAAL 100UG/DOSIS | N02AB03 | Fentanyl nasaal | **Fentanyl oromucosaal** | **verkeerd** (M-03), wel oranje |
| FENTANYL INJVLST 50UG/ML | N02AB03 | Fentanyl s.c./i.v. | idem | goed |
| HYDROMORFON CAPSULE MGA 4MG | N02AA03 | Hydromorfon oraal | idem | goed |
| HYDROMORFON CAPSULE MGA 4MG | (leeg) | Hydromorfon oraal | **Morfine oraal** | **stil verkeerd** (H-04) |
| PALLADON SLOW 8MG | (leeg) | Hydromorfon oraal | geen opioïd | **stil weggelaten** (M-04a) |
| TRAMADOL/PARACETAMOL TABLET 37,5/325MG | N02AJ13 | Tramadol 37,5 | 37,5 | goed |
| PARACETAMOL/CODEINE TABLET 500/20MG | N02AJ06 | Codeïne 20 | 20 | goed |
| CODEINE TABLET 10MG | R05DA04 | Codeïne 10 | 10 | goed |
| METHADON DRANK 1MG/ML | N07BC02 | Methadon 1 mg/ml | idem | goed, oranje |
| TAPENTADOL TABLET MGA 50MG | N02AX06 | Tapentadol 50 | 50 | goed |
| OXYCONTIN 10 | N02AA05 | Oxycodon 10 | sterkte niet herkend | goed (oranje, zelf invullen) |
| OXYCODON HCL INJVLST 10MG/ML | N02AA05 | Oxycodon s.c./i.v. | idem | goed |
| MORFINE TABLET MGA 30MG | (leeg) | Morfine oraal 30 | idem | goed (op naam) |

**End-to-end (`Verwerk`)**: V01 (oxycodon 2.1T + P1CM6 + fentanyl 12 D31PL, eGFR 45) gaf exact 40 / 30 (doorbraak, oranje) / 12, en eGFR 45 in D22. eGFR `<15` → 15, oranje, met nierwaarschuwing (goed). eGFR `29,5` → 29,5 met waarschuwing (goed). eGFR `45 (CKD-EPI 2009)` → 452009 (M-05). Een einddatum met tweecijferig jaartal werkt hier door de Windows-instelling (L-11).

---

## Koppeling en PowerShell (stap 4), alleen gelezen

- **Toetsen**: `{F3}` (maximaal 2× bij "niet gevonden"), daarna `%i`, `%d` en `{F3}`, alleen als Pharmacom het voorgrondvenster is en het hoofdvenster *enabled* is (geen modale melding). Dat is een goede beveiliging. Er blijft een klein tijdvenster tussen controle en `SendWait` (L-07).
- **Verkeerde patiënt**: de eGFR wordt alleen overgenomen als de venstertitel van het lab-uitlezen gelijk is aan die van de medicatiestatus (`Pharmacom_naar_Excel.ps1:124`). Goed, mits de titel de patiënt identificeert (**vraag**: staat naam of geboortedatum in de titel?). B4 toont de titel met "controleer of dit de juiste patiënt is": goed.
- **Race**: wisselt de gebruiker van dossier tussen het lezen van de medicatie en het lezen van het lab, dan voorkomt de titelvergelijking een verkeerde eGFR. De medicatie zelf wordt in één keer gelezen.
- **Foutafhandeling**: het leesscript zet elke fout om in `FOUT<tab>reden` (`trap`). `Verwerk` heeft `On Error`. Ontbrekende kolommen worden gemeld. Netjes.
- **Encoding**: UTF-8 zonder BOM van PS naar VBA. De VBA bevat bewust alleen ASCII (via `T()`), en de build zet de `.bas` om naar cp1252. µ wordt in de sterkte genormaliseerd naar `U`. Getest met `Codeïne` en `µg`: goed.

## Privacy en beveiliging (stap 5)

- Er worden geen gegevens naar het netwerk gestuurd, er zijn geen logbestanden, en de VBA bevat geen `Shell` (het uitlezen start bewust buiten Excel). Dat is een goed ontwerp voor omgevingen met macrobeperkingen.
- Tijdelijk bestand: zie M-08. Gegevens in de werkmap of AutoHerstel: zie M-07. `.gitignore` sluit `opioidrotatie_*.txt` uit (goed).
- `Installeer.ps1`: schrijft alleen in het gebruikersprofiel (Downloads, Bureaublad, `%USERPROFILE%\.accessibility.properties`), leest HKLM alleen, downloadt niets en vraagt geen beheerdersrechten. `-ExecutionPolicy Bypass` in de snelkoppeling is nodig en beperkt tot dit ene script. De Java Access Bridge-instelling geldt voor alle Java-programma's van de gebruiker (wordt gemeld).
- AVG: het programma verwerkt gezondheidsgegevens lokaal. Voor gebruik in een apotheekorganisatie: laten toetsen door de FG/privacyfunctionaris (verwerking binnen het eigen AIS-proces; vermoedelijk geen aparte DPIA nodig, maar dat is **te laten beoordelen**).

## Kwaliteit en onderhoudbaarheid (stap 6)

- `werkmap_opbouw.py` (1275 r.) is leesbaar en consistent. Celadressen worden grotendeels afgeleid (`DET`, `R_*`, `W0`), wat goed is. Het zwakke punt is de koppeling met de VBA via losse constanten en commentaar (de oorzaak van K-03). Advies: benoemde bereiken.
- Magische getallen staan bijna allemaal als instelling op tabblad *Omrekenfactoren* (goed). Uitzonderingen: de score `0.03*(E-1)` en de maxima `E>4` / `E>2` in *Omzetten*, en de 90 dagen voor "eGFR oud".
- Elke factor heeft een bronvermelding. Een raadpleegdatum per factor ontbreekt (L-10).
- README ↔ werkmap: de factoren in de README-tabel kloppen allemaal met de werkmap (geteste voorbeelden: 60 mg → 20/40/20/12/4/25/25/300/150/1200 µg/200 µg/15). Piritramide en codeïne ontbreken in de README (L-05).

## Testen en CI (stap 7)

- `bron/Test_werkmap.ps1` dekt codes, sterktes, `Verwerk` en *Omzetten* breed, maar **zonder één assertie** (M-09). Het detecteert dus niets automatisch.
- Nieuw in `review/`:
  - `referentie.py`: onafhankelijke berekening (pomp, methadonstaffel heen en terug, preparaatkeuze "dichtstbij" en "veilig").
  - `parser_port.py`: getrouwe port van de VBA-parser.
  - `casussen.py` + `excel_harness.ps1`: casussen door de echte werkmap, uitkomst in `werkmap_uitkomsten.json`.
  - `test_referentie.py` (pytest): **235 geslaagd, 21 xfail**. Elke bekende bevinding is een `xfail(strict=True)`: wordt de fout opgelost, dan faalt de test totdat de markering weg is. Inclusief een **layouttest** die de VBA-constanten (`R_EERST`, `CEL_EGFR`, het `EigenCode`-bereik) controleert tegen een verse openpyxl-build. Die draait op Linux en had K-03 voorkomen.
- `ci-voorstel.yml`: een Linux-job (pytest, layout) plus een self-hosted Windows-job met Excel (bouwen, harness, pytest). GitHub-hosted runners hebben geen Office.

## Regelgeving (stap 8): te laten beoordelen, geen juridisch advies

- **EU-MDR 2017/745**: software die bedoeld is om informatie te geven voor therapeutische beslissingen valt onder **regel 11** (in beginsel klasse IIa of hoger). Het beoogde doel bepaalt de kwalificatie, niet de disclaimer. Een rekenhulp die een patiëntspecifieke pompstand berekent, ligt dicht bij die definitie. De README-zin "geen medisch hulpmiddel" dekt dat niet af. De zin "ondersteunt maar vervangt geen klinisch oordeel / dubbele controle" is wel waardevol voor het gebruik.
- **In-house-uitzondering (art. 5 lid 5 MDR)**: die kan gelden voor gebruik binnen één zorginstelling, onder voorwaarden (onder meer een kwaliteitssysteem, onderbouwing en een openbare verklaring). Of een openbare apotheek(keten) hieronder valt, en wat publieke verspreiding via GitHub (MIT) betekent voor "in de handel brengen", moet worden **beoordeeld door een deskundige** (bijv. de IGJ-informatiepunten of een MDR-adviseur).
- **Beroepsrichtlijnen**: KNMP en NVZA hebben eisen of handreikingen voor zelfgebouwde rekenhulpen en validatie. Ik heb de actuele documenten **niet geverifieerd**; laat de toepasselijkheid toetsen. Een minimaal validatiedossier zou bevatten: beoogd gebruik, bronnen per factor met datum, testplan met verwachte uitkomsten (de casussen hier zijn een begin), vrijgave door een tweede apotheker, versiebeheer en wijzigingslog.

---

## Aanbevelingen in volgorde

**1. Patiëntveiligheid**
| # | Aanbeveling | Bevinding | Moeite |
|---|---|---|---|
| 1 | Doorbraakadvies nooit boven de berekende dosis, plus de >10%-melding voor doorbraak | K-01 | S |
| 2 | Pompstand naar beneden afronden, of geen stand tonen bij >10% afwijking | K-02, M-02 | S |
| 3 | `EigenCode` naar `B9:F58` (benoemd bereik) + layouttest | K-03 | S |
| 4 | Geen groen advies boven +10%, en geen preparaat boven de max. dagdosis | H-01, H-02 | M |
| 5 | Validatie op de eGFR-cel en tekst als `<15` herkennen | H-03 | S |
| 6 | `HYDROMORF` vóór `MORF`; herkenning op naam altijd oranje | H-04 | S |
| 7 | Uitzoeken wat *Actief* betekent en niet-actieve en dubbele regels uitsluiten of melden | H-05 | M |
| 8 | Elke doseercodescheider verifiëren tegen de Centrale doseertabel, onzekere oranje | H-06 | M |

**2. Quick wins (< 1 uur)**
| # | Aanbeveling | Bevinding | Moeite |
|---|---|---|---|
| 9 | `NASAAL` toevoegen; nasaal vóór `SPRAY` | M-03 | S |
| 10 | Negatieve dosis en tekst in het dosisveld in de Controle-kolom | M-01, L-04 | S |
| 11 | eGFR: eerste getal nemen, >200 weigeren | M-05 | S |
| 12 | `try/finally` om het tijdelijke bestand | M-08 | S |
| 13 | Spuit-melding, lege-invoermelding, buprenorfine-opmerking, README-tabel en VBA-kopcommentaar | L-01…L-05 | S |
| 14 | Regels zonder ATC-code altijd in de samenvatting | M-04 | S |

**3. Structureel**
| # | Aanbeveling | Bevinding | Moeite |
|---|---|---|---|
| 15 | `Workbook_BeforeClose`/`BeforeSave` wissen/waarschuwen, AutoHerstel uit | M-07 | M |
| 16 | Benoemde bereiken als enige koppeling Python ↔ VBA | L-06 | M |
| 17 | Asserties in `Test_werkmap.ps1`, of de `review/`-tests overnemen in `tests/` + CI | M-09 | M |
| 18 | Raadpleegdatum per factor; sterktes 12,5/37,5 controleren | L-10 | S |
| 19 | Validatiedossier + MDR/beroepsrichtlijn-toets | M-10 | L |

## Open vragen voor de auteur

1. Toont de medicatiestatus (F3) alleen actuele medicatie, of ook gestopte, toekomstige of dubbele regels (herhaling)? Wat staat er in *Actief*? (H-05)
2. Wat betekenen de scheiders `U`, `S` en `D` na de frequentie in de Centrale doseertabel (bijv. `4U1T`)? (H-06)
3. Staat de patiëntnaam of -geboortedatum in de venstertitel van Pharmacom? (De controle op hetzelfde dossier en B4 hangen daarvan af.)
4. Komt de eGFR in Pharmacom ooit met tekst of voetnoten in de cel (bijv. `45 (CKD-EPI 2009)`, `45*`)? (M-05)
5. Rectaal wordt 1:1 als oraal gerekend: is dat bewust en volgens welke bron?
6. Fentanylpleister "12" en "37": worden die bedoeld als 12,5 en 37,5 µg/u?
7. Is afronden naar de dichtstbijzijnde pompstap (in plaats van naar beneden) een bewuste keuze, bijv. uit een lokaal protocol? (K-02)
8. Waar wordt de rekenhulp gebruikt (één apotheek, keten, ziekenhuis)? Dat is relevant voor de MDR-beoordeling.

---

## Wat is uitgevoerd en wat niet

**Uitgevoerd**
- De werkmap opnieuw gebouwd met `werkmap_opbouw.py` (openpyxl) en vergeleken met de `.xlsm` in de repo: 0 formuleverschillen.
- 35 pompcasussen, 14 omzetcasussen en 1 casus met leeg assortiment door de **echte werkmap** in Excel (COM, alleen-lezen, niet opgeslagen).
- 37 doseercodes en 30 etiketnamen door de **echte VBA** (`VertaalCodeTekst`, `SterkteTekst`), 11 end-to-end-scenario's via `Verwerk` (verzonnen gegevens), en een test van eigen doseercodes.
- Onafhankelijke referentie en parser-port in Python. `pytest`: 235 geslaagd, 21 xfail (bekende bevindingen).

**Alleen gelezen (niet uitgevoerd)**
- `Pharmacom_medicatiestatus.ps1` en `Pharmacom_naar_Excel.ps1` zonder Pharmacom: de echte Java Access Bridge, SendKeys en de Pharmacom-schermen zijn niet getest. Er is geen Pharmacom aangesproken.
- `Installeer.ps1` is niet gedraaid (dat zou snelkoppelingen en `.accessibility.properties` wijzigen).
- `Bouw_werkmap.ps1` niet gedraaid (de bestaande `.xlsm` is gelijk aan de bron, zie boven). `Test_werkmap.ps1` niet apart gedraaid; de harness dekt dezelfde functies en meer.
- LibreOffice is hier niet geïnstalleerd. Dat was ook niet nodig, omdat Excel zelf beschikbaar was.
- KNMP-factoren zijn **niet** tegen de actuele KNMP-tekst geverifieerd (geen toegang); alleen intern vergeleken (README ↔ code ↔ werkmap). De betekenis van doseercodes is niet tegen de officiële Centrale doseertabel geverifieerd.
