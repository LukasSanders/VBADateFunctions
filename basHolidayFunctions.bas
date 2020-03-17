Attribute VB_Name = "basHolidayFunctions"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ### Funktionsbausteine ###
'
' Modul:    basHolidayFunctions
' Autor:    Lukas Sanders
' Datum:    14.10.2018
' Version:  1.0
' Lizenz:   (C) 2018 Lukas Sanders. Nutzung und Wiederverwendung unter MIT-Lizenz.
' Funktion: Beispieldatei zur Vorstellung von Hilfsprozeduren zur Datenaufbereitung
'           in UserForms
'
' Dieses Modul enth�lt verschiedene Funktionen, welche Feiertage in der BRD
' errechnen bzw. zur�ckgeben sowie eine Funktion zur Pr�fung, ob ein Datum
' ein Feiertag ist.
'
'
'
' Erforderliche Dateien/Verkn�pfungen:
' ------------------------------------
' Die Funktionen f�r variable Feiertagsdaten erfordern die Funktionen f�r den
' Ostersonntag bzw. den 1. Adventssonntag.
'
' Ebenfalls erfordert die Funktion IsHoliday alle Feiertagsfunktionen.
' Setzen Sie dieses Modul somit bevorzugt vollst�ndig ein!
'
' Bei folgenden Ereignissen m�ssen Funktionen angepasst bzw. erweitert werden:
'
' - bei (einmaligen) zus�tzlichen Feiertagen, wobei der Reformationstag,
'   welcher 2017 einmalig bundesweit Feiertag war, bereits ber�cksichtigt wird;
'   hier w�ren eine neue Feiertagsfunktion anzulegen und die Funktion
'   IsHoliday entsprechend zu erg�nzen
' - falls die Bev�lkerung von Augsburg nicht mehr �berwiegend katholisch sein
'   sollte und daher Mari� Himmelfahrt dort kein Feiertag ist; hier w�re die
'   Funktion IsHoliday an der entsprechenden Stelle anzupassen
'
' Kompatibilit�t:
' ---------------
' MS Excel ab 2013 (getestet)
'
'
' Bekannte Probleme:
' ------------------
' Da Excel annimmt, dass das Jahr 1900 ein Schaltjahr ist, kann es in diesem Jahr zu
' Fehlern oder unerw�nschtem Verhalten kommen.
'
'
' Funktionen und Prozeduren:
' --------------------------
'
' # Allgemein Feiertage         Alle Funktionen haben das Jahr als Eingabewert (Integer)
'                               und geben das Datum des jeweiligen Feiertags als Date zur�ck
'   Eingabewerte:
'       year (erforderlich)     Jahr als Integer
'   R�ckgabewerte:
'       [Funktionsname]         Date
'
' Feste Feiertage:
' ----------------
' # Heiligabend                 Heiligabend (24.12.)
' # ErsterWeihnachtstag         Erster Weihnachtsfeiertag (25.12.)
' # ZweiterWeihnachtstag        Zweiter Weihnachtsfeiertag (26.12.)
' # Sylvester                   Sylvester (31.12.)
' # Neujahr                     Neujahr (01.01.)
' # DreiKoenigsTag              Heilige Drei K�nige (06.01.)
' # TagDerArbeit                Tag der Arbeit (01.05.)
' # Friedensfest                Augsburger Friedensfest (08.08.)
' # MariaeHimmelfahrt           Mari� Himmelfahrt (15.08.)
' # TagDerDtEinheit             Tag der Deutschen Einheit (03.10.)
' # Allerheiligen               Allerheiligen (01.11.)
' # Reformationstag             Reformationstag (31.10.)
'
'
' Variable Feiertage:
' -------------------
' # OsterSonntag                Ostersonntag
' # ErsterAdvent                1. Advent
'
'
' Variable Feiertage (abh�ngig):
' ------------------------------
' # Aschermittwoch              Aschermittwoch, 46 Tage vor Ostersonntag
' # Karfreitag                  Karfreitag, 2 Tage vor Ostersonntag
' # Karsamstag                  Karsamstag, 1 Tag vor Ostersonntag
' # OsterMontag                 Ostermontag, 1 Tag nach Ostersonntag
' # ChristiHimmelfahrt          Christi Himmelfahrt, 39 Tage nach Ostersonntag
' # PfingstSonntag              Pfingstsonntag, 49 Tage nach Ostersonntag
' # PfingstMontag               Pfingstmontag, 50 Tage nach Ostersonntag
' # Fronleichnam                Fronleichnam, 60 Tage nach Ostersonntag
' # BussUndBettag               Bu�- und Bettag, 11 Tage vor dem 1. Advent
' # ZweiterAdvent               2. Advent, 7 Tage nach dem 1. Advent
' # DritterAdvent               3. Advent, 14 Tage nach dem 1. Advent
' # VierterAdvent               4. Advent, 21 Tage nach dem 1. Advent
'
' # IsHoliday                   Gibt zur�ck, ob ein Datum ein Feiertag ist
'       Eingabewerte:
'           datum               Zu �berpr�fendes Datum (Date)
'           bland               K�rzel Bundesland (String)
'       R�ckgabewerte:
'           IsHoliday           Boolean
'       Fehlernummern:
'           20                  Bundesland nicht definiert
'           21                  sonstiger Fehler, ggf. ung�ltiges Datum eingegeben
'
'       Wertetabelle bland:
'           BY                  Bayern
'           BX                  Bayern (Augsburg)
'           BZ                  Bayern (Gemeinden mit �berwiegend kath. Bev�lkerung)
'           BW                  Baden-W�rttemberg
'           BE                  Berlin
'           BB                  Brandenburg
'           HB                  Hansestadt Bremen
'           HH                  Hansestadt Hamburg
'           HE                  Hessen
'           MV                  Mecklemburg-Vorpommern
'           NI                  Niedersachsen
'           NW                  Nordrhein-Westfalen
'           RP                  Rheinland-Pfalz
'           SL                  Saarland
'           SN                  Sachsen
'           SX                  Sachsen (Gemeinden, in denen Fronleichnam lt. Verordnung Feiertag ist)
'           ST                  Sachsen-Anhalt
'           SH                  Schleswig-Holstein
'           TH                  Th�ringen
'           TX                  Th�ringen (Gemeinden, in denen Fronleichnam lt. Verordnung Feiertag ist)
'           BU                  nur bundeseinheitliche Feiertage
'
'           Systemk�rzel BX wird ben�tigt, da das Friedensfest nur in Augsburg ein Feiertag ist.
'           Systemk�rzel BZ wird ben�tigt, da Mari� Himmelfahrt in Bayern nur in Gemeinden mit
'               �berwiegend katholischer Bev�lkerung ein Feiertag ist; Augsburg z�hlt (derzeit)
'               zu diesen Gemeinden.
'           Systemk�rzel SX wird ben�tigt, da Fronleichnam in Sachsen nur in Gemeinden mit
'               �berwiegend katholischer Bev�lkerung lt. Rechtsverordnung ein Feiertag ist
'           Systemk�rzel TX wird ben�tigt, da Fronleichnam in Th�ringen nur in durch
'               Rechtsverordnung bestimmten Gemeinden oder aufgrund Fortgeltung
'               alten Rechts ein Feiertag ist
'           Systemk�rzel BU wurde hinzugef�gt, um explizit nur bundeseinheitliche Feiertage
'               abfragen zu k�nnen.
'
' # ListHolidays                Listet alle Feiertage f�r ein bestimmtes Jahr sowie ein bestimmtes
'                               Bundesland im Arbeitsblatt unterhalb der aktivierten Zelle auf
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Funktionen f�r feste Feiertagsdaten
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Heiligabend(year As Integer) As Date
    Heiligabend = DateSerial(year, 12, 24)
End Function

Public Function ErsterWeihnachtstag(year As Integer) As Date
    ErsterWeihnachtstag = DateSerial(year, 12, 25)
End Function

Public Function ZweiterWeihnachtstag(year As Integer) As Date
    ZweiterWeihnachtstag = DateSerial(year, 12, 26)
End Function

Public Function Sylvester(year As Integer) As Date
    Sylvester = DateSerial(year, 12, 31)
End Function

Public Function Neujahr(year As Integer) As Date
    Neujahr = DateSerial(year, 1, 1)
End Function

Public Function DreiKoenigsTag(year As Integer) As Date
    DreiKoenigsTag = DateSerial(year, 1, 6)
End Function

Public Function TagDerArbeit(year As Integer) As Date
    TagDerArbeit = DateSerial(year, 5, 1)
End Function

Public Function Friedensfest(year As Integer) As Date
    Friedensfest = DateSerial(year, 8, 8)
End Function

Public Function MariaeHimmelfahrt(year As Integer) As Date
    MariaeHimmelfahrt = DateSerial(year, 8, 15)
End Function

Public Function TagDerDtEinheit(year As Integer) As Date
    TagDerDtEinheit = DateSerial(year, 10, 3)
End Function

Public Function Allerheiligen(year As Integer) As Date
    Allerheiligen = DateSerial(year, 11, 1)
End Function

Public Function Reformationstag(year As Integer) As Date
    Reformationstag = DateSerial(year, 10, 31)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Funktionen f�r die Erreichnung des 1. Adventssonntags und des Ostersonntags
'
' Achtung:
' --------
' Diese Daten werden f�r die nachfolgenden Feiertagsdaten als Rechengrundlage ben�tigt!
' Die nachfolgenden Funktionen sind somit zwingend erforderlich f�r die Berechnung
' der variablen Feiertagsdaten!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function OsterSonntag(year As Integer) As Date
' Umsetzung von Spencers Osterformel, beschrieben 1922 von Harold Spencer Jones

    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, _
        h As Long, i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, _
        o As Long, p As Long
    
    a = year Mod 19
    
    b = WorksheetFunction.Quotient(year, 100)
        
    c = year Mod 100
    
    d = WorksheetFunction.Quotient(b, 4)
    
    e = b Mod 4
    
    f = WorksheetFunction.Quotient(b + 8, 25)
    
    g = WorksheetFunction.Quotient(b - f + 1, 3)
    
    h = (19 * a + b - d - g + 15) Mod 30
    
    i = WorksheetFunction.Quotient(c, 4)
    
    k = c Mod 4
    
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    
    m = WorksheetFunction.Quotient(a + 11 * h + 22 * l, 451)
    
    n = WorksheetFunction.Quotient(h + l - 7 * m + 114, 31)
    
    o = (h + l - 7 * m + 114) Mod 31
    
    p = o + 1
    
    OsterSonntag = DateSerial(year, n, p)
End Function

Public Function ErsterAdvent(year As Integer) As Date
' Zun�chst auf den letzten Sonntag vor dem 1. Weihnachtsfeiertag zur�ckgehen (4. Advent),
' dann 3 Wochen zur�ckrechnen
'
' letzter Sonntag vor dem 25.12. wird, wenn Sonntag der 7. Wochentag ist , durch Subtraktion der
' Wochentagsnummer des 25.12. errechnet

    ErsterAdvent = DateSerial(year, 12, 25) - Weekday(DateSerial(year, 12, 25), vbMonday) - 21
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Funktionen f�r �brige variable Feiertagsdaten
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Aschermittwoch(year As Integer) As Date
    Aschermittwoch = OsterSonntag(year) - 46
End Function

Public Function Karfreitag(year As Integer) As Date
    Karfreitag = OsterSonntag(year) - 2
End Function

Public Function Karsamstag(year As Integer) As Date
    Karsamstag = OsterSonntag(year) - 1
End Function

Public Function OsterMontag(year As Integer) As Date
    OsterMontag = OsterSonntag(year) + 1
End Function

Public Function ChristiHimmelfahrt(year As Integer) As Date
    ChristiHimmelfahrt = OsterSonntag(year) + 39
End Function

Public Function PfingstSonntag(year As Integer) As Date
    PfingstSonntag = OsterSonntag(year) + 49
End Function

Public Function PfingstMontag(year As Integer) As Date
    PfingstMontag = OsterSonntag(year) + 50
End Function

Public Function Fronleichnam(year As Integer) As Date
    Fronleichnam = OsterSonntag(year) + 60
End Function

Public Function BussUndBettag(year As Integer) As Date
    BussUndBettag = ErsterAdvent(year) - 11
End Function

Public Function ZweiterAdvent(year As Integer) As Date
    ZweiterAdvent = ErsterAdvent(year) + 7
End Function

Public Function DritterAdvent(year As Integer) As Date
    DritterAdvent = ErsterAdvent(year) + 14
End Function

Public Function VierterAdvent(year As Integer) As Date
    VierterAdvent = ErsterAdvent(year) + 21
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Funktionen f�r die Pr�fung, ob ein Feiertag vorliegt
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function IsHoliday(datum As Date, bland As String) As Boolean
' �berpr�ft, ob ein bestimmtes �bergebenes Datum ein Feiertag ist
' R�ckgabe je nach Bundesland

On Error GoTo Fehler

' R�ckgabewert standardm��ig Falsch
IsHoliday = False


' ########### Eingegebenes Bundeslandk�rzel �berpr�fen

' Zun�chst �berpr�fen, ob ein g�ltiges Bundesland angegeben wurde
' (zul�ssige Werte siehe Modulbeschreibung)

Dim arrBLand
Dim i As Integer
Dim bolIsValidBLand As Boolean

arrBLand = Array("BY", "BX", "BZ", "BW", "BU", "BE", "BB", "HB", "HH", "HE", "MV", "NI", "NW", "RP", "SL", "SN", _
        "ST", "SH", "TH", "SX", "TX")

For i = 0 To UBound(arrBLand)
    If arrBLand(i) = bland Then
        bolIsValidBLand = True
        Exit For
    Else
        bolIsValidBLand = False
    End If
Next i

' Falls das Bundesland in der Liste nicht gefunden wurde, Fehler ausgeben und abbrechen
If bolIsValidBLand = False Then
    Err.Raise Number:=20, Description:="Angegebenes Bundesland ung�ltig!"
    Exit Function
End If


' ########### Bundeseinheitliche Feiertage �berpr�fen

' F�r bundeseinheitliche Feiertage muss das Bundesland-K�rzel nicht ausgelesen werden

If datum = Neujahr(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = Karfreitag(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = OsterSonntag(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = OsterMontag(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = TagDerArbeit(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = ChristiHimmelfahrt(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = PfingstSonntag(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = PfingstMontag(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = TagDerDtEinheit(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = ErsterWeihnachtstag(year(datum)) Then
    IsHoliday = True
    Exit Function
End If

If datum = ZweiterWeihnachtstag(year(datum)) Then
    IsHoliday = True
    Exit Function
End If



' ########### L�nderspezifische Feiertage �berpr�fen

' F�r l�nderspezifische Feiertage muss das Bundesland-K�rzel ausgelesen werden

' Dreik�nigstag ist nur in Bayern und Baden-W�rttemberg ein gesetzlicher Feiertag.
'
' Bayern (BY) schlie�t hier die Systemk�rzel BX (Augsburg) und BZ (kath. Gemeinden)
' mit ein.

If bland = "BW" Or bland = "BY" Or bland = "BX" Or bland = "BZ" Then
    If datum = DreiKoenigsTag(year(datum)) Then
        IsHoliday = True
        Exit Function
    End If
End If

' Fronleichnam ist in Bayern, Baden-W�rttemberg, Hamburg, NRW, Rheinland-Pfalz und dem
' Saarland immer ein gesetzlicher Feiertag, zus�tzlich
'   - in Sachsen in den  vom Staatsministerium des Inneren durch Rechtsverordnung bestimmten Gemeinden
'     im Landkreis Bautzen und im Westlausitzkreis (Systemk�rzel SX) sowie
'   - in Th�ringen in durch Rechtsverordnung f�r Gemeinden mit �berwiegend katholischer Bev�lkerung oder
'     in Gemeinden, in denen bis 1994 Fronleichnam als gesetzlicher Feiertag begangen wurde,
'     bis zum Erlass einer solchen Rechtsverordnung.
'
' Bayern (BY) schlie�t hier die Systemk�rzel BX (Augsburg) und BZ (kath. Gemeinden)
' mit ein.

If bland = "BW" Or bland = "BY" Or bland = "BX" Or bland = "BZ" Or bland = "HH" Or bland = "NW" Or bland = "RP" _
        Or bland = "SX" Or bland = "TX" Then
    If datum = Fronleichnam(year(datum)) Then
        IsHoliday = True
        Exit Function
    End If
End If

' Das Augsburger Friedensfest ist nur in Augsburg ein gesetzlicher Feiertag (Systemk�rzel BX).

If bland = "BX" Then
    If datum = Friedensfest(year(datum)) Then
        IsHoliday = True
        Exit Function
    End If
End If

' Mari� Himmelfahrt ist nur im Saarland sowie in Gemeinden in Bayern mit �berwiegend katholischer Bev�lkerung
' ein gesetzlicher Feiertag (Systemk�rzel BZ), hierzu z�hlt (derzeit) auch Augsburg (Systemk�rzel BX).
'
' Falls die Bev�lkerung in Augsburg irgendwann nicht mehr �berwiegend katholisch sein sollte, was vergleichsweise
' unwahrscheinlich ist, m�sste diese Funktion angepasst werden.

If bland = "BX" Or bland = "BZ" Or bland = "SL" Then
    If datum = MariaeHimmelfahrt(year(datum)) Then
        IsHoliday = True
        Exit Function
    End If
End If

' Folgenden Block aktivieren und vorherigen Block auskommentieren, falls die Augsburger Bev�lkerung
' nicht mehr �berwiegend katholisch sein sollte.

'If bland = "BZ" Or bland = "SL" Then
'    If datum = MariaeHimmelfahrt(year(datum)) Then
'        IsHoliday = True
'        Exit Function
'    End If
'End If

' Reformationstag ist nur in Berlin, Bremen, Hamburg, Mecklemburg-Vorpommern, Niedersachsen, Sachsen,
' Sachsen-Anhalt, Schleswig-Holstein und Th�ringen ein gesetzlicher Feiertag.
'
' Im Jahr 2017 war der Reformationstag jedoch zum 500. Jahrestags einmalig ein bundesweiter Feiertag.

If bland = "BB" Or bland = "HB" Or bland = "HH" Or bland = "MV" Or bland = "NI" Or bland = "SN" _
        Or bland = "ST" Or bland = "SH" Or bland = "TH" Or bland = "SX" Or bland = "TX" _
        Or year(datum) = 2017 Then
    If datum = Reformationstag(year(datum)) Then
        IsHoliday = True
        Exit Function
    End If
End If

' Der Bu�- und Bettag ist nur in Sachsen ein gesetzlicher Feiertag.
'
' Sachsen (SN) Schlie�t hier das Systemk�rzel SX mit ein.

If bland = "SN" Or bland = "SX" Then
    If datum = BussUndBettag(year(datum)) Then
        IsHoliday = True
        Exit Function
    End If
End If

Exit Function

' Allerheiligen ist nur in Bayern, Baden-W�rttemberg, NRW, Rheinland-Pfalz
' und dem Saarland ein gesetzlicher Feiertag.
'
' Bayern (BY) Schlie�t hier die Systemk�rzel BX und BZ mit ein.

If bland = "BY" Or bland = "BX" Or bland = "BZ" Or bland = "BW" Or bland = "NW" _
        Or bland = "RP" Or bland = "SL" Then
    If datum = Allerheiligen(year(datum)) Then
        IsHoliday = True
        Exit Function
    End If
End If

Exit Function


' ########### Fehlerbehandlung f�r �brige Fehler
Fehler:
Err.Raise Number:=21, Description:="Fehler aufgetreten, ggf. Datum ung�ltig."

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Prozedur zur Auflistung aller Feiertage eines Jahres
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub ListHolidays()
' F�gt an der aktuellen Stelle eine Liste der Feiertage des aktuellen Jahres ein

On Error GoTo Fehler

' Aktuelle Zelle (Startzelle) auslesen
Dim intCurRow As Integer
Dim intCurCol As Integer

intCurRow = ActiveCell.Row
intCurCol = ActiveCell.Column

' Jahr und L�nderk�rzel abfragen
Dim intYear As Integer
Dim strBLand As String

intYear = InputBox("Bitte geben Sie das Jahr ein, f�r welches die Feiertage " _
            & "aufgelistet werden sollen!", "Bitte Jahr eingeben")

' Zun�chst pr�fen, ob ein g�ltiges Jahr eingegeben wurde
'
' Wenn die Eingabe nicht numerisch ist, "knallt" es schon bei der Wert�bergabe und die Fehlerbehandlung
' wird aufgerufen, deshalb muss hier nur noch der zul�ssige Wertebereich �berpr�ft werden.
'
' Mit Werten vor 1900 kommt Excel nicht so gut zurecht, deshalb lassen wir nur 1900 - 9999 zu.
If intYear < 1900 Or intYear > 9999 Then
    MsgBox "Angegebenes Jahr ung�ltig!", vbCritical + vbOKOnly, "Fehler - Jahreseingabe ung�ltig"
    Exit Sub
End If

' Anmerkung: Die folgende Abfrage w�re sicherlich einfacher und ansehnlicher �ber ein UserForm mit einem
' Listenfeld zu l�sen, jedoch ist diese Prozedur als Makro f�r den Anwender sinnvoller auszuf�hren und zu
' integrieren. Ein UserForm w�rde letztlich die Einbindung zu stark verkomplizieren.

strBLand = InputBox("Bitte geben Sie das Bundesland an!" & vbCrLf & vbCrLf & vbCrLf _
            & "K�rzel:" & vbCrLf & vbCrLf _
            & "BY Bayern" & vbCrLf _
            & "BX Bayern(Augsburg)" & vbCrLf _
            & "BZ Bayern (mit Mari� Himmelfahrt)" & vbCrLf _
            & "BW Baden - W�rttemberg" & vbCrLf _
            & "BE Berlin" & vbCrLf _
            & "BB Brandenburg" & vbCrLf _
            & "HB Hansestadt Bremen" & vbCrLf _
            & "HH Hansestadt Hamburg" & vbCrLf _
            & "HE Hessen" & vbCrLf _
            & "MV Mecklemburg - Vorpommern" & vbCrLf _
            & "NI Niedersachsen" & vbCrLf _
            & "NW Nordrhein - Westfalen" & vbCrLf _
            & "RP Rheinland - Pfalz" & vbCrLf _
            & "SL Saarland" & vbCrLf _
            & "SN Sachsen" & vbCrLf _
            & "SX Sachsen (mit Fronleichnam)" & vbCrLf _
            & "ST Sachsen - Anhalt" & vbCrLf _
            & "SH Schleswig - Holstein" & vbCrLf _
            & "TH Th�ringen" & vbCrLf _
            & "TX Th�ringen (mit Fronleichnam)" & vbCrLf _
            & "BU nur bundeseinheitliche Feiertage" & vbCrLf, _
            "Bitte Bundesland angeben")
            
' Zun�chst pr�fen, ob ein g�ltiges Bundesland angegeben wurde
Dim arrBLand
Dim i As Integer
Dim bolIsValidBLand As Boolean

arrBLand = Array("BY", "BX", "BZ", "BW", "BU", "BE", "BB", "HB", "HH", "HE", "MV", "NI", "NW", "RP", "SL", "SN", _
        "ST", "SH", "TH", "SX", "TX")

For i = 0 To UBound(arrBLand)
    If arrBLand(i) = strBLand Then
        bolIsValidBLand = True
        Exit For
    Else
        bolIsValidBLand = False
    End If
Next i

' Falls das Bundesland in der Liste nicht gefunden wurde, Fehler ausgeben und abbrechen
If bolIsValidBLand = False Then
    MsgBox "Angegebenes Bundesland ung�ltig!", vbCritical + vbOKOnly, "Fehler - Bundeslandk�rzel ung�ltig"
    Exit Sub
End If

' Ansonsten Tage des Jahres durchlaufen und f�r jeden Tag �berpr�fen, ob ein Feiertag vorliegt
Dim datStartDatum As Date
Dim j As Integer
Dim k As Integer

datStartDatum = DateSerial(intYear, 1, 1)

If IsLeapYear(intYear) = True Then
    k = 365
Else
    k = 364
End If

For j = 0 To k
    If IsHoliday(datStartDatum + j, strBLand) = True Then
        ' Wenn das aktuelle Datum ein Feiertag ist, wird das Datum in die aktuelle Zelle geschrieben
        Cells(intCurRow, intCurCol).Value = datStartDatum + j
        ' Die Zellennummer wird um 1 erh�ht, damit beim n�chsten Treffer die n�chste Zelle ge�ndert wird
        intCurRow = intCurRow + 1
    End If
Next j

Exit Sub

' Fehlerbehandlung f�r sonstige Fehler
Fehler:
    MsgBox "Eingabe ung�ltig!", vbCritical + vbOKOnly, "Fehler - Eingabe ung�ltig"
    Exit Sub
End Sub
