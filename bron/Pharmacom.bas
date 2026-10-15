Attribute VB_Name = "Pharmacom"
' =====================================================================
' Opioidrotatie: huidige opioiden uitlezen uit Pharmacom (medicatiestatus, F3)
'
' Het uitlezen start NIET vanuit Excel: op veel pc's mogen macro's in Office geen
' PowerShell starten (beveiliging; Excel wordt dan afgesloten of de start geweigerd).
' De snelkoppeling "Lees opioiden uit Pharmacom" (sneltoets Ctrl+Alt+F3) start
' Pharmacom_naar_Excel.ps1; dat leest via de Java Access Bridge de medicatiestatus
' en roept hier Verwerk aan met het resultaat.
' Herkent opioiden aan de ATC-code, de sterkte aan de etiketnaam en vertaalt
' de doseercode volgens de systematiek van de Centrale doseertabel (Health
' Base): frequentie - tijdsaanduiding - hoeveelheid - toedieningsvorm, met
' eventueel "M" + getal = niet meer dan ... per dag.
' Codes op tabblad Doseercodes gaan voor de automatische vertaling.
' Alleen lezen: er worden geen toetsen naar Pharmacom gestuurd.
' (Geen niet-ASCII tekens in deze module; zie T() voor e-trema, i-trema, mu en de pijl.)
' =====================================================================
Option Explicit

Private Const R_EERST As Long = 10       ' blok 1 (zie werkmap_opbouw.py: R0..R1)
Private Const R_LAATST As Long = 15
Private Const KLEUR_INVOER As Long = 16777215   ' wit (invulveld)
Private Const KLEUR_LET_OP As Long = 15069695   ' RGB(255, 241, 229), licht BENU-oranje

Private Type Dosering
    Herkend As Boolean
    ZoNodig As Boolean
    PerDagMin As Double      ' eenheden per 24 u
    PerDagMax As Double      ' -1 = onbekend
    MaxPerDag As Double      ' "niet meer dan", -1 = geen
    Aantal As Double         ' eenheden per keer
    AantalMax As Double
    IntervalDagen As Double  ' 1 keer per n dagen, 0 = n.v.t.
    Tijden As Boolean        ' code met tijdsaanduidingen (ochtend, avond, ...)
    Vorm As String
    Tekst As String
End Type

Private Type Sterkte
    Gevonden As Boolean
    Waarde As Double
    Eenheid As String        ' MG of UG
    Per As String            ' "", ML of UUR
    PerAantal As Double      ' bijv. 5 bij 10MG/5ML
    Opmerking As String
End Type

' ---------------------------------------------------------------- knoppen
Public Sub LeesUitPharmacom()
    MsgBox "Uitlezen gaat via de sneltoets Ctrl+Alt+F3 (of de snelkoppeling 'Lees opioiden uit Pharmacom' op het bureaublad):" & vbLf & vbLf & _
           "1. Open in Pharmacom het dossier van de patient." & vbLf & _
           "2. Druk op Ctrl+Alt+F3 (mag ook vanuit Excel): Pharmacom wordt dan zelf gekozen en zo nodig de medicatiestatus (F3) geopend." & vbLf & vbLf & _
           "Blok 1 wordt dan hier ingevuld. (Excel zelf mag op veel computers geen programma's starten.)", _
           vbInformation, "Lees uit Pharmacom"
End Sub

Public Sub WisInvoer()
    If MsgBox("Blok 1 en de uitgelezen gegevens wissen?", vbYesNo + vbQuestion, "Wis invoer") = vbNo Then Exit Sub
    WisBlok1
    ThisWorkbook.Worksheets("Rotatie").Range("B4").Value = Plaatshouder()
End Sub

Private Function Plaatshouder() As String
    Plaatshouder = "Nog niet uitgelezen uit Pharmacom  " & ChrW(183) & "  open het dossier en druk op Ctrl+Alt+F3"
End Function

Private Sub WisBlok1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Rotatie")
    ws.Range("B" & R_EERST & ":D" & R_LAATST).ClearContents
    ws.Range("B" & R_EERST & ":D" & R_LAATST).Interior.Color = KLEUR_INVOER
    ws.Range("I" & R_EERST & ":I" & R_LAATST).ClearContents
    ws.Range("I" & R_EERST & ":I" & R_LAATST).Interior.Pattern = xlNone
    Dim c As Range
    For Each c In ws.Range("B" & R_EERST & ":D" & R_LAATST)
        If Not c.Comment Is Nothing Then c.Comment.Delete
    Next c
End Sub

' ---------------------------------------------------------------- verwerken
' Verwerkt de uitvoer van het uitleesscript (ook los te testen met verzonnen regels).
Public Function Verwerk(ByVal Tekst As String, Optional ByVal Stil As Boolean = False) As String
    On Error GoTo Fout
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Rotatie")

    If Not Stil And Application.WorksheetFunction.CountA(ws.Range("B" & R_EERST & ":D" & R_LAATST)) > 0 Then
        If MsgBox("Blok 1 bevat al gegevens. Overschrijven met de gegevens uit Pharmacom?", vbYesNo + vbQuestion, "Lees uit Pharmacom") = vbNo Then Exit Function
    End If

    Dim Regels() As String
    Regels = Split(Replace(Tekst, vbCr, ""), vbLf)
    Dim Kop() As String
    Kop = Split(Regels(0), vbTab)
    If Kop(0) <> "OK" Then
        Verwerk = Melding(Stil, "Uitlezen lukte niet: " & IIf(UBound(Kop) >= 1, Kop(UBound(Kop)), Regels(0)), vbExclamation)
        Exit Function
    End If
    Dim Titel As String
    Titel = Trim$(Replace(Kop(1), "Pharmacom - ", ""))
    If InStr(1, Titel, "Dossier", vbTextCompare) = 0 Then
        Verwerk = Melding(Stil, "In Pharmacom staat geen dossier open (venster: " & Titel & ").", vbExclamation)
        Exit Function
    End If

    ' Kolomnummers uit de kopregel
    Dim H() As String, iNaam As Long, iDos As Long, iEind As Long, iAtc As Long, iHoev As Long, j As Long
    H = Split(Regels(1), vbTab)
    iNaam = -1: iDos = -1: iEind = -1: iAtc = -1: iHoev = -1
    For j = 0 To UBound(H)
        Select Case H(j)
            Case "Etiketnaam": iNaam = j
            Case "Dosering": iDos = j
            Case "Th. einddatum": iEind = j
            Case "ATC": iAtc = j
            Case "Afl. Hoev.": iHoev = j
        End Select
    Next j
    If iNaam < 0 Or iDos < 0 Then
        Verwerk = Melding(Stil, "De medicatiestatus had niet de verwachte kolommen (Etiketnaam, Dosering).", vbExclamation)
        Exit Function
    End If

    WisBlok1
    ws.Range("B4").Value = T("Uitgelezen uit Pharmacom: ") & Titel & " (" & Format(Now, "dd-mm-yyyy hh:nn") & T(")  {->}  controleer of dit de juiste pati{e}nt is.")

    Dim r As Long, n As Long, Overgeslagen As String, NietIngevuld As String, Velden() As String
    r = R_EERST
    For n = 2 To UBound(Regels)
        If Trim$(Regels(n)) <> "" Then
            Velden = Split(Regels(n), vbTab)
            Dim Naam As String, Code As String, Atc As String, Eind As String
            Naam = Veld(Velden, iNaam): Code = Veld(Velden, iDos)
            Atc = Veld(Velden, iAtc): Eind = Veld(Velden, iEind)
            Dim Stof As String
            Stof = Opioid(Atc, Naam)
            If Stof <> "" Then
                If IsDatum(Eind) Then
                    If DatumVan(Eind) < Date Then
                        Overgeslagen = Overgeslagen & vbLf & "- " & Naam & " (" & Code & T("): therapie be{e}indigd op ") & Eind
                        GoTo Volgende
                    End If
                End If
                If r > R_LAATST Then
                    NietIngevuld = NietIngevuld & vbLf & "- " & Naam & " (" & Code & T("): geen ruimte meer in blok 1")
                    GoTo Volgende
                End If
                Dim Opm As String, Kort As String
                If VulRegel(ws, r, Stof, Naam, Code, Veld(Velden, iHoev), Opm) Then
                    r = r + 1
                Else
                    NietIngevuld = NietIngevuld & vbLf & "- " & Naam & " (" & Code & "): " & Opm
                End If
            End If
        End If
Volgende:
    Next n

    Dim Aantal As Long
    Aantal = r - R_EERST
    Dim Samenvatting As String
    Samenvatting = Aantal & T(" opio{i}dregel(s) ingevuld uit ") & Titel & "."
    If Aantal > 0 Then Samenvatting = Samenvatting & vbLf & "Controleer de kolom 'Uitgelezen uit Pharmacom' (oranje = nakijken / aanvullen)."
    If NietIngevuld <> "" Then Samenvatting = Samenvatting & vbLf & vbLf & "Niet automatisch ingevuld (zelf beoordelen):" & NietIngevuld
    If Overgeslagen <> "" Then Samenvatting = Samenvatting & vbLf & vbLf & "Overgeslagen:" & Overgeslagen
    If Aantal = 0 And NietIngevuld = "" Then Samenvatting = T("Geen opio{i}den gevonden in de medicatiestatus van ") & Titel & "."
    ws.Activate
    Verwerk = Melding(Stil, Samenvatting, vbInformation)
    Exit Function
Fout:
    Application.StatusBar = False
    Verwerk = Melding(Stil, "Onverwachte fout bij het verwerken: " & Err.Description & " (" & Err.Number & ")", vbCritical)
End Function

' Vult regel r van blok 1. Geeft False (met reden in Opm) als de regel niet
' ingevuld kan worden.
Private Function VulRegel(ws As Worksheet, ByVal r As Long, ByVal Stof As String, ByVal Naam As String, _
                          ByVal Code As String, ByVal Hoev As String, ByRef Opm As String) As Boolean
    Dim Weg As String, Middel As String
    Weg = Toedieningsweg(Naam)
    Middel = TabelNaam(Stof, Weg)
    If Middel = "" Then
        Opm = Stof & " " & Weg & T(" staat niet in de omrekentabel (KNMP: geen omrekenfactor)")
        Exit Function
    End If

    Dim D As Dosering, S As Sterkte
    D = VertaalCode(Code)
    S = LeesSterkte(Naam, Stof)

    Dim Soort As String, Dosis As Variant, LetOp As Boolean, Uitleg As String
    Soort = "Onderhoud": Dosis = Empty
    Uitleg = Naam & "  |  " & Code & " = " & IIf(D.Herkend, D.Tekst, "code niet herkend")

    If Not S.Gevonden Then
        LetOp = True
        Uitleg = Uitleg & T("  {->}  sterkte niet herkend in de etiketnaam; vul de dosis zelf in.")
    ElseIf Right$(Middel, 8) = "pleister" Then
        ' Pleister: dosis = sterkte (ug/uur) x aantal tegelijk geplakte pleisters
        If S.Per <> "UUR" Then
            LetOp = True
            Uitleg = Uitleg & T("  {->}  sterkte per uur niet herkend; vul {mu}g/uur zelf in.")
        Else
            Dim Pl As Double
            Pl = 1
            If D.Herkend And D.Aantal > 0 Then Pl = D.Aantal
            Dosis = S.Waarde * Pl * IIf(S.Eenheid = "MG", 1000, 1)
            Uitleg = Uitleg & "  " & T("{->}  ") & Getal(Dosis) & T(" {mu}g/uur")
            If Not D.Herkend Then
                LetOp = True
                Uitleg = Uitleg & " (aangenomen: 1 pleister)"
            ElseIf D.AantalMax > D.Aantal Then
                LetOp = True
                Uitleg = Uitleg & " (bereik " & Getal(D.Aantal) & "-" & Getal(D.AantalMax) & " pleisters: controleer)"
            End If
        End If
    Else
        ' mg (of ug) per eenheid
        Dim PerEenheid As Double, EenheidTekst As String
        PerEenheid = S.Waarde
        If S.Per = "ML" And S.PerAantal > 0 Then PerEenheid = PerEenheid / S.PerAantal
        EenheidTekst = IIf(S.Eenheid = "UG", T("{mu}g"), "mg")
        If S.Per = "ML" And D.Herkend And D.Vorm <> "ML" And D.Vorm <> "M" And D.Vorm <> "" Then
            LetOp = True
            Uitleg = Uitleg & T("  {->}  let op: sterkte per ml maar vorm '") & D.Vorm & "'"
        End If
        If S.Opmerking <> "" Then LetOp = True: Uitleg = Uitleg & "  (" & S.Opmerking & ")"

        If Not D.Herkend Then
            LetOp = True
            Uitleg = Uitleg & T("  {->}  vul de dosis per 24 u zelf in (of voeg de code toe op tabblad Doseercodes).")
        ElseIf D.ZoNodig Then
            Soort = "Doorbraak"
            LetOp = True
            ' Zo nodig / bij pijn: standaard het maximum per 24 u (de gebruiker verlaagt zo nodig naar het werkelijke gebruik)
            If D.PerDagMax > 0 Then
                Dosis = D.PerDagMax * PerEenheid
                Uitleg = Uitleg & T("  {->}  zo nodig/bij pijn: MAXIMUM ingevuld (") & Getal(Dosis) & " " & EenheidTekst & _
                         "/24 u). Verlaag naar het werkelijke gebruik als dat minder is."
            Else
                Uitleg = Uitleg & T("  {->}  zo nodig/bij pijn zonder maximum in de code: vul het werkelijke gebruik van de afgelopen 24 u in.")
            End If
        Else
            Dosis = D.PerDagMin * PerEenheid
            Uitleg = Uitleg & T("  {->}  ") & Getal(Dosis) & " " & EenheidTekst & "/24 u"
            If D.PerDagMax > D.PerDagMin Then
                LetOp = True
                Uitleg = Uitleg & " (bereik " & Getal(D.PerDagMin * PerEenheid) & "-" & Getal(D.PerDagMax * PerEenheid) & _
                         " " & EenheidTekst & "/24 u; laagste ingevuld - controleer werkelijk gebruik)"
            End If
        End If

        ' Continue fentanyl s.c./i.v. rekent in ug/uur
        If Middel = "Fentanyl s.c./i.v. continu" And Not IsEmpty(Dosis) Then
            If S.Eenheid = "MG" Then Dosis = Dosis * 1000
            Dosis = Dosis / 24
            Uitleg = Uitleg & T(" = ") & Getal(Dosis) & T(" {mu}g/uur")
        End If
        ' Snelwerkend fentanyl (oromucosaal/nasaal) rekent in ug/24 u; factor indicatief
        If Middel = "Fentanyl oromucosaal" Or Middel = "Fentanyl nasaal" Then
            LetOp = True
            If Not IsEmpty(Dosis) And S.Eenheid = "MG" Then
                Dosis = Dosis * 1000
                Uitleg = Uitleg & T(" = ") & Getal(Dosis) & T(" {mu}g/24 u")
            End If
        End If
        If Middel = "Methadon oraal" Then LetOp = True
    End If

    ws.Cells(r, "B").Value = Soort
    ws.Cells(r, "C").Value = Middel
    If Not IsEmpty(Dosis) Then ws.Cells(r, "D").Value = Round(Dosis, 2)
    ws.Cells(r, "I").Value = Uitleg
    If LetOp Then
        ws.Cells(r, "I").Interior.Color = KLEUR_LET_OP
        If IsEmpty(Dosis) Then ws.Cells(r, "D").Interior.Color = KLEUR_LET_OP
    End If
    VulRegel = True
End Function

' ---------------------------------------------------------------- middel
' Werkzame stof bij een opioid, of "" als het geen opioid is.
Private Function Opioid(ByVal Atc As String, ByVal Naam As String) As String
    Dim A As String
    A = UCase$(Trim$(Atc))
    Select Case Left$(A, 7)
        Case "N02AA01": Opioid = "Morfine"
        Case "N02AA05", "N02AA55", "N02AA56": Opioid = "Oxycodon"
        Case "N02AA03": Opioid = "Hydromorfon"
        Case "N02AB03": Opioid = "Fentanyl"
        Case "N02AE01", "N07BC01", "N07BC51": Opioid = "Buprenorfine"
        Case "N02AX02", "N02AJ13", "N02AJ14", "N02AJ15": Opioid = "Tramadol"
        Case "N02AX06": Opioid = "Tapentadol"
        Case "N02AC03": Opioid = "Piritramide"
        Case "N07BC02", "N02AC52": Opioid = "Methadon"
        Case "R05DA04", "N02AA59", "N02AA79", "N02AJ06", "N02AJ07", "N02AJ09": Opioid = T("Code{i}ne")
    End Select
    If Opioid <> "" Then Exit Function
    If A <> "" And Left$(A, 4) <> "N02A" And Left$(A, 5) <> "N07BC" Then Exit Function
    ' Geen (bekende) ATC-code: op naam
    Dim N As String, i As Long, Stoffen As Variant
    N = UCase$(Naam)
    Stoffen = Array("MORF", "Morfine", "OXYCOD", "Oxycodon", "HYDROMORF", "Hydromorfon", "FENTAN", "Fentanyl", _
                    "BUPREN", "Buprenorfine", "TRAMAD", "Tramadol", "TAPENT", "Tapentadol", "PIRITR", "Piritramide", _
                    "METHAD", "Methadon", "CODEI", T("Code{i}ne"))
    For i = 0 To UBound(Stoffen) Step 2
        If InStr(N, Stoffen(i)) > 0 Then Opioid = Stoffen(i + 1): Exit Function
    Next i
    If Left$(A, 4) = "N02A" Then Opioid = T("Onbekend opio{i}d")
End Function

Private Function Zoekwoord(ByVal Stof As String) As String
    Select Case Stof
        Case "Morfine": Zoekwoord = "MORF"
        Case "Oxycodon": Zoekwoord = "OXYCOD"
        Case "Hydromorfon": Zoekwoord = "HYDROMORF"
        Case "Fentanyl": Zoekwoord = "FENTAN"
        Case "Buprenorfine": Zoekwoord = "BUPREN"
        Case "Tramadol": Zoekwoord = "TRAMAD"
        Case "Tapentadol": Zoekwoord = "TAPENT"
        Case "Piritramide": Zoekwoord = "PIRITR"
        Case "Methadon": Zoekwoord = "METHAD"
        Case Else: Zoekwoord = "CODEI"
    End Select
End Function

' oraal, parenteraal, pleister, nasaal, transmucosaal of rectaal (uit de etiketnaam)
Private Function Toedieningsweg(ByVal Naam As String) As String
    Dim N As String
    N = " " & UCase$(Naam) & " "
    If BevatEen(N, Array("PLEISTER", "PLAKKER", " TTS", "MATRIX")) Then
        Toedieningsweg = "pleister"
    ElseIf BevatEen(N, Array("INJ", "INF", "AMPUL", "CASSETTE", " IV ", " SC ")) Then
        Toedieningsweg = "parenteraal"
    ElseIf BevatEen(N, Array("NEUS", "NASAL", "INSTANYL", "PECFENT")) Then
        Toedieningsweg = "nasaal"
    ElseIf BevatEen(N, Array("SPRAY", "ZUIGT", "SUBL", "BUCC", "FILM", "OROMUC", "SMELT", "LOLLY", _
                             "ACTIQ", "ABSTRAL", "EFFENTORA", "BREAKYL")) Then
        Toedieningsweg = "transmucosaal"
    ElseIf BevatEen(N, Array("ZETPIL", "SUPP", "RECT")) Then
        Toedieningsweg = "rectaal"
    Else
        Toedieningsweg = "oraal"
    End If
End Function

' Naam zoals op tabblad Omrekenfactoren, of "" als er geen omrekening is.
Private Function TabelNaam(ByVal Stof As String, ByVal Weg As String) As String
    Dim Oraal As Boolean, Par As Boolean
    Oraal = (Weg = "oraal" Or Weg = "rectaal")
    Par = (Weg = "parenteraal")
    Select Case Stof
        Case "Morfine", "Oxycodon", "Hydromorfon"
            If Oraal Then TabelNaam = Stof & " oraal"
            If Par Then TabelNaam = Stof & " s.c./i.v."
        Case "Fentanyl"
            If Weg = "pleister" Then TabelNaam = "Fentanyl pleister"
            If Par Then TabelNaam = "Fentanyl s.c./i.v. continu"
            If Weg = "transmucosaal" Then TabelNaam = "Fentanyl oromucosaal"
            If Weg = "nasaal" Then TabelNaam = "Fentanyl nasaal"
        Case "Buprenorfine"
            If Weg = "pleister" Then TabelNaam = "Buprenorfine pleister"
        Case "Tramadol", "Tapentadol", "Methadon"
            If Oraal Then TabelNaam = Stof & " oraal"
        Case "Piritramide"
            If Par Then TabelNaam = "Piritramide s.c./i.v."
        Case T("Code{i}ne")
            If Oraal Then TabelNaam = Stof & " oraal"
    End Select
End Function

' ---------------------------------------------------------------- sterkte
' Eerste getal + MG/UG in de etiketnaam, met eventueel /ML of /UUR erachter.
' Bij combinatiepreparaten (bijv. PARACETAMOL/CODEINE 500/10MG) het getal op
' de plaats van het opioid.
Private Function LeesSterkte(ByVal Naam As String, ByVal Stof As String) As Sterkte
    Dim S As Sterkte, N As String, i As Long, k As Long, Tok As String, Rest As String, Eenh As String
    N = UCase$(Replace(Replace(Naam, ChrW(181), "U"), ChrW(956), "U"))
    i = 1
    Do While i <= Len(N)
        Dim Begin As Boolean
        Begin = IsCijfer(Mid$(N, i, 1))
        If Begin And i > 1 Then Begin = Not IsLetterOfCijfer(Mid$(N, i - 1, 1))
        If Begin Then
            k = i
            Do While k <= Len(N)
                If InStr("0123456789.,/", Mid$(N, k, 1)) = 0 Then Exit Do
                k = k + 1
            Loop
            Tok = Mid$(N, i, k - i)
            Rest = LTrim$(Mid$(N, k))
            Eenh = ""
            If Left$(Rest, 9) = "MICROGRAM" Then
                Eenh = "UG": Rest = Mid$(Rest, 10)
            ElseIf Left$(Rest, 3) = "MCG" Then
                Eenh = "UG": Rest = Mid$(Rest, 4)
            ElseIf Left$(Rest, 2) = "UG" Then
                Eenh = "UG": Rest = Mid$(Rest, 3)
            ElseIf Left$(Rest, 2) = "MG" Then
                Eenh = "MG": Rest = Mid$(Rest, 3)
            End If
            If Eenh <> "" Then
                Dim Getallen() As String, Idx As Long, Delen() As String, Ingr As String
                Getallen = Split(Tok, "/")
                Idx = 0
                Ingr = Left$(N, i - 1)
                If InStr(Ingr, "/") > 0 Then
                    Delen = Split(Ingr, "/")
                    For k = 0 To UBound(Delen)
                        If InStr(Delen(k), Zoekwoord(Stof)) > 0 Then Idx = k: Exit For
                    Next k
                    If UBound(Getallen) <> UBound(Delen) Then S.Opmerking = "combinatiepreparaat: controleer sterkte"
                ElseIf UBound(Getallen) > 0 Then
                    S.Opmerking = "meerdere sterktes in naam: controleer"
                End If
                If Idx > UBound(Getallen) Then Idx = 0
                S.Waarde = Getal2(Getallen(Idx))
                S.Eenheid = Eenh
                S.Gevonden = S.Waarde > 0
                ' /ML, /5ML, /UUR, /U, /H, /24UUR
                Rest = LTrim$(Rest)
                If Left$(Rest, 1) = "/" Then
                    Rest = LTrim$(Mid$(Rest, 2))
                    k = 1
                    Do While k <= Len(Rest)
                        If InStr("0123456789.,", Mid$(Rest, k, 1)) = 0 Then Exit Do
                        k = k + 1
                    Loop
                    If k > 1 Then S.PerAantal = Getal2(Left$(Rest, k - 1))
                    Rest = LTrim$(Mid$(Rest, k))
                    If Left$(Rest, 2) = "ML" Then
                        S.Per = "ML"
                    ElseIf Left$(Rest, 1) = "U" Or Left$(Rest, 1) = "H" Then
                        If S.PerAantal = 0 Or S.PerAantal = 1 Then S.Per = "UUR"
                    End If
                End If
                LeesSterkte = S
                Exit Function
            End If
            i = k
        Else
            i = i + 1
        End If
    Loop
    LeesSterkte = S
End Function

' ---------------------------------------------------------------- doseercode
Private Function VertaalCode(ByVal Code As String) As Dosering
    Dim D As Dosering, C As String, i As Long, P As Long
    D.MaxPerDag = -1: D.PerDagMax = -1
    C = UCase$(Replace(Trim$(Code), " ", ""))
    If C = "" Then VertaalCode = D: Exit Function

    ' 1. Eigen tabel (tabblad Doseercodes)
    If EigenCode(C, D) Then VertaalCode = D: Exit Function

    Dim Kern As String, Vorm As String
    Kern = C
    ' "niet meer dan": M + getal aan het eind
    P = InStrRev(Kern, "M")
    If P > 2 And P < Len(Kern) Then
        If IsGetal(Mid$(Kern, P + 1)) Then
            D.MaxPerDag = Val(Mid$(Kern, P + 1))
            Kern = Left$(Kern, P - 1)
        End If
    End If
    ' Toedieningsvorm (positie 4+5)
    Dim Vormen As Variant
    Vormen = Array("PL", "ZK", "ML", "IJ", "C", "T", "S", "M")
    For i = 0 To UBound(Vormen)
        If Len(Kern) > Len(Vormen(i)) And Right$(Kern, Len(Vormen(i))) = Vormen(i) Then
            ' "M" alleen als milliliter na een hoeveelheid (anders is het "ochtend")
            If Not (Vormen(i) = "M" And Not IsHoeveelheidTeken(Mid$(Kern, Len(Kern) - 1, 1))) Then
                Vorm = Vormen(i)
                Kern = Left$(Kern, Len(Kern) - Len(Vormen(i)))
                Exit For
            End If
        End If
    Next i
    D.Vorm = Vorm

    Dim Fmin As Double, Fmax As Double, Amin As Double, Amax As Double, Sep As String, Rest As String
    Dim Ok As Boolean, Periode As Double
    Periode = 1

    If InStr(Kern, "-") > 0 Then
        ' Vaste dag in de week, bijv. ZO-1 / MA-1: 1 keer per week
        Rest = Mid$(Kern, InStr(Kern, "-") + 1)
        If Len(Left$(Kern, InStr(Kern, "-") - 1)) = 2 And Hoeveelheid(Rest, Amin, Amax) Then
            Fmin = 1: Fmax = 1: Periode = 7: Ok = True
        End If
    ElseIf Len(Kern) >= 3 And IsCijfer(Left$(Kern, 1)) Then
        ' Frequentie (1 of 2 cijfers) + . / W / P / Z / U + hoeveelheid
        Fmin = Val(Left$(Kern, 1)): Fmax = Fmin
        i = 2
        If IsCijfer(Mid$(Kern, 2, 1)) And InStr(".WPZUSD", Mid$(Kern, 3, 1)) > 0 And Len(Kern) >= 4 Then
            Fmax = Val(Mid$(Kern, 2, 1)): i = 3
        End If
        Sep = Mid$(Kern, i, 1)
        If InStr(".WPZUSD", Sep) > 0 And Sep <> "" Then
            If Hoeveelheid(Mid$(Kern, i + 1), Amin, Amax) Then
                Ok = True
                If Sep = "W" Then Periode = 7
                If Sep = "P" Or Sep = "Z" Then D.ZoNodig = True
            End If
        End If
    ElseIf (Left$(Kern, 1) = "P" Or Left$(Kern, 1) = "Z") And Len(Kern) >= 2 Then
        ' Bij pijn / zo nodig + hoeveelheid (frequentie niet vastgelegd)
        If Hoeveelheid(Mid$(Kern, 2), Amin, Amax) Then
            Ok = True: D.ZoNodig = True: Fmin = 0: Fmax = -1
        End If
    ElseIf Left$(Kern, 1) = "D" And Len(Kern) >= 3 And IsCijfer(Mid$(Kern, 2, 1)) Then
        ' 1 keer per n dagen
        If Hoeveelheid(Mid$(Kern, 3), Amin, Amax) Then
            Ok = True: Fmin = 1: Fmax = 1: Periode = Val(Mid$(Kern, 2, 1)): D.IntervalDagen = Periode
        End If
    ElseIf Left$(Kern, 2) = "OD" Then
        If Hoeveelheid(Mid$(Kern, 3), Amin, Amax) Then
            Ok = True: Fmin = 1: Fmax = 1: Periode = 2: D.IntervalDagen = 2
        End If
    Else
        Ok = Tijden(Kern, Amin, Amax, D.ZoNodig)
        Fmin = 1: Fmax = 1: D.Tijden = True
    End If

    If Not Ok Then VertaalCode = D: Exit Function
    D.Herkend = True
    D.Aantal = Amin: D.AantalMax = Amax
    D.PerDagMin = Fmin * Amin / Periode
    If Fmax >= 0 Then D.PerDagMax = Fmax * Amax / Periode
    If D.ZoNodig Then
        D.PerDagMin = 0
        If D.MaxPerDag > 0 And (D.PerDagMax < 0 Or D.MaxPerDag < D.PerDagMax) Then D.PerDagMax = D.MaxPerDag
    End If
    D.Tekst = Omschrijving(D, Fmin, Fmax, Periode, Amin, Amax)
    VertaalCode = D
End Function

' Tijdsaanduidingen: M(ochtend) MI(middag) A(vond) N(acht) 8 @ ( met elk
' eventueel P/Z en een hoeveelheid; een tijd zonder hoeveelheid krijgt die van
' de volgende tijd (MA1 = ochtend 1, avond 1). Telt op tot eenheden per dag.
Private Function Tijden(ByVal K As String, ByRef Amin As Double, ByRef Amax As Double, ByRef ZoNodig As Boolean) As Boolean
    Dim i As Long, Wacht As Long, Tok As String, Mn As Double, Mx As Double, Aantal As Long
    i = 1
    Do While i <= Len(K)
        If Mid$(K, i, 2) = "MI" Then
            i = i + 2
        ElseIf InStr("MAN8@(", Mid$(K, i, 1)) > 0 Then
            i = i + 1
        Else
            Exit Function
        End If
        Wacht = Wacht + 1
        If InStr("PZ", Mid$(K, i, 1)) > 0 And i <= Len(K) Then ZoNodig = True: i = i + 1
        Tok = ""
        Do While i <= Len(K) And Len(Tok) < 2
            If InStr("0123456789KHQ)%", Mid$(K, i, 1)) = 0 Then Exit Do
            Tok = Tok & Mid$(K, i, 1)
            i = i + 1
        Loop
        If Tok <> "" Then
            If Not Hoeveelheid(Tok, Mn, Mx) Then Exit Function
            Amin = Amin + Wacht * Mn: Amax = Amax + Wacht * Mx
            Aantal = Aantal + Wacht
            Wacht = 0
        End If
    Loop
    Tijden = (Aantal > 0 And Wacht = 0)
End Function

' Hoeveelheid (positie 3): 1-9, K, H, Q, A, ) = 10, % = 15; twee tekens = bereik (12 = 1 tot 2, H1 = halve tot 1)
Private Function Hoeveelheid(ByVal S As String, ByRef Mn As Double, ByRef Mx As Double) As Boolean
    Dim A As Double, B As Double
    If Len(S) = 1 Then
        A = Teken(S)
        If A <= 0 Then Exit Function
        Mn = A: Mx = A: Hoeveelheid = True
    ElseIf Len(S) = 2 Then
        A = Teken(Left$(S, 1)): B = Teken(Right$(S, 1))
        If A <= 0 Or B <= 0 Or B <= A Then Exit Function
        Mn = A: Mx = B: Hoeveelheid = True
    End If
End Function

Private Function Teken(ByVal C As String) As Double
    Select Case C
        Case "1" To "9": Teken = Val(C)
        Case "K": Teken = 0.25
        Case "H": Teken = 0.5
        Case "Q": Teken = 0.75
        Case "A": Teken = 1.5
        Case ")": Teken = 10
        Case "%": Teken = 15
        Case Else: Teken = -1
    End Select
End Function

Private Function IsHoeveelheidTeken(ByVal C As String) As Boolean
    IsHoeveelheidTeken = (Teken(C) > 0)
End Function

Private Function Omschrijving(D As Dosering, ByVal Fmin As Double, ByVal Fmax As Double, ByVal Periode As Double, _
                              ByVal Amin As Double, ByVal Amax As Double) As String
    Dim S As String, Vorm As String
    Select Case D.Vorm
        Case "C": Vorm = "capsule"
        Case "T": Vorm = "tablet"
        Case "PL": Vorm = "pleister"
        Case "ZK": Vorm = "zakje"
        Case "ML", "M": Vorm = "ml"
        Case "S": Vorm = "zetpil"
        Case "IJ": Vorm = "injectie"
        Case Else: Vorm = "eenheid"
    End Select
    If D.ZoNodig Then S = "zo nodig/bij pijn "
    If D.Tijden Then
        S = S & "verdeeld over de dag, in totaal "
    ElseIf Fmax >= 0 And Fmin > 0 Then
        S = S & Getal(Fmin) & IIf(Fmax > Fmin, " tot " & Getal(Fmax), "") & " keer per "
        If Periode = 1 Then
            S = S & "dag "
        ElseIf Periode = 7 Then
            S = S & "week "
        Else
            S = S & Getal(Periode) & " dagen "
        End If
    End If
    S = S & Getal(Amin) & IIf(Amax > Amin, " tot " & Getal(Amax), "") & " " & Vorm
    If D.MaxPerDag > 0 Then S = S & ", niet meer dan " & Getal(D.MaxPerDag) & " per dag"
    If Not D.ZoNodig Then S = S & " (= " & Getal(D.PerDagMin) & IIf(D.PerDagMax > D.PerDagMin, "-" & Getal(D.PerDagMax), "") & "/24 u)"
    Omschrijving = S
End Function

' Code op tabblad Doseercodes (A7:E56)
Private Function EigenCode(ByVal C As String, ByRef D As Dosering) As Boolean
    Dim Rij As Range, ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Doseercodes")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For Each Rij In ws.Range("A7:A56").Cells
        If UCase$(Replace(Trim$(CStr(Rij.Value)), " ", "")) = C And C <> "" Then
            D.Herkend = True
            D.ZoNodig = (UCase$(Trim$(CStr(Rij.Offset(0, 2).Value))) = "J")
            If IsNumeric(Rij.Offset(0, 1).Value) And Not IsEmpty(Rij.Offset(0, 1).Value) Then
                D.PerDagMin = CDbl(Rij.Offset(0, 1).Value): D.PerDagMax = D.PerDagMin
                D.Aantal = D.PerDagMin: D.AantalMax = D.PerDagMin
            ElseIf Not D.ZoNodig Then
                D.Herkend = False
            End If
            If IsNumeric(Rij.Offset(0, 3).Value) And Not IsEmpty(Rij.Offset(0, 3).Value) Then
                D.MaxPerDag = CDbl(Rij.Offset(0, 3).Value)
                If D.ZoNodig Then D.PerDagMax = D.MaxPerDag
            End If
            If D.ZoNodig Then D.PerDagMin = 0
            D.Tekst = Trim$(CStr(Rij.Offset(0, 4).Value)) & " (eigen tabel)"
            EigenCode = True
            Exit Function
        End If
    Next Rij
End Function

' ---------------------------------------------------------------- testen
' Voor controle: =VertaalCodeTekst("P1CM6") of via Application.Run
Public Function VertaalCodeTekst(ByVal Code As String) As String
    On Error GoTo Fout
    Dim D As Dosering
    D = VertaalCode(Code)
    If Not D.Herkend Then
        VertaalCodeTekst = "niet herkend"
    Else
        VertaalCodeTekst = D.Tekst & " | min " & D.PerDagMin & " | max " & D.PerDagMax & " | zo nodig " & D.ZoNodig & " | aantal " & D.Aantal & " | vorm " & D.Vorm
    End If
    Exit Function
Fout:
    VertaalCodeTekst = "FOUT: " & Err.Description
End Function

Public Function SterkteTekst(ByVal Naam As String, ByVal Atc As String) As String
    On Error GoTo Fout
    Dim S As Sterkte, Stof As String
    Stof = Opioid(Atc, Naam)
    If Stof = "" Then SterkteTekst = "geen opioid": Exit Function
    S = LeesSterkte(Naam, Stof)
    SterkteTekst = Stof & " | " & Toedieningsweg(Naam) & " | " & TabelNaam(Stof, Toedieningsweg(Naam)) & " | " & _
                   IIf(S.Gevonden, S.Waarde & " " & S.Eenheid & IIf(S.Per <> "", "/" & IIf(S.PerAantal > 1, S.PerAantal, "") & S.Per, ""), "sterkte?") & _
                   IIf(S.Opmerking <> "", " | " & S.Opmerking, "")
    Exit Function
Fout:
    SterkteTekst = "FOUT: " & Err.Description
End Function

' ---------------------------------------------------------------- hulpfuncties
Private Function Melding(ByVal Stil As Boolean, ByVal Tekst As String, ByVal Soort As VbMsgBoxStyle) As String
    Melding = Tekst
    If Not Stil Then MsgBox Tekst, Soort, "Lees uit Pharmacom"
End Function

Private Function LeesUtf8(ByVal Pad As String) As String
    Dim St As Object
    Set St = CreateObject("ADODB.Stream")
    St.Type = 2
    St.Charset = "utf-8"
    St.Open
    St.LoadFromFile Pad
    LeesUtf8 = St.ReadText
    St.Close
End Function

Private Function Veld(V() As String, ByVal i As Long) As String
    If i >= 0 And i <= UBound(V) Then Veld = Trim$(V(i))
End Function

Private Function IsDatum(ByVal S As String) As Boolean
    Dim P() As String
    P = Split(S, "-")
    If UBound(P) = 2 Then IsDatum = IsGetal(P(0)) And IsGetal(P(1)) And IsGetal(P(2))
End Function

Private Function DatumVan(ByVal S As String) As Date
    Dim P() As String
    P = Split(S, "-")
    DatumVan = DateSerial(CInt(P(2)), CInt(P(1)), CInt(P(0)))
End Function

Private Function IsGetal(ByVal S As String) As Boolean
    Dim i As Long
    If S = "" Then Exit Function
    For i = 1 To Len(S)
        If Not IsCijfer(Mid$(S, i, 1)) Then Exit Function
    Next i
    IsGetal = True
End Function

Private Function IsCijfer(ByVal C As String) As Boolean
    IsCijfer = (C >= "0" And C <= "9" And Len(C) = 1)
End Function

Private Function IsLetterOfCijfer(ByVal C As String) As Boolean
    IsLetterOfCijfer = IsCijfer(C) Or (UCase$(C) >= "A" And UCase$(C) <= "Z")
End Function

Private Function BevatEen(ByVal Tekst As String, ByVal Woorden As Variant) As Boolean
    Dim i As Long
    For i = 0 To UBound(Woorden)
        If InStr(Tekst, Woorden(i)) > 0 Then BevatEen = True: Exit Function
    Next i
End Function

' "37,5" / "37.5" -> 37.5 (onafhankelijk van landinstellingen)
Private Function Getal2(ByVal S As String) As Double
    Getal2 = Val(Replace(S, ",", "."))
End Function

Private Function Getal(ByVal X As Double) As String
    If X = Int(X) Then
        Getal = CStr(X)
    Else
        Getal = Format$(X, "0.0##")
    End If
End Function

Private Function T(ByVal S As String) As String
    S = Replace(S, "{e}", ChrW(235))
    S = Replace(S, "{i}", ChrW(239))
    S = Replace(S, "{mu}", ChrW(181))
    S = Replace(S, "{->}", ChrW(8594))
    T = S
End Function
