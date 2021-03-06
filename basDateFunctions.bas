Attribute VB_Name = "basDateFunctions"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ### Funktionsbausteine ###
'
' Modul:    basDateFunctions
' Autor:    Lukas Sanders
' Datum:    07.10.2018
' Version:  1.0
' Lizenz:   (C) 2018 Lukas Sanders. Nutzung und Wiederverwendung unter MIT-Lizenz.
' Funktion: Beispieldatei zur Vorstellung von Hilfsprozeduren zur Datenaufbereitung
'           in UserForms
'
' Dieses Modul enth�lt verschiedene Funktionen, welche h�ufig ben�tigte Prozedur-
' schritte bei der Verarbeitung von Datumswertenvereinfachen sollen.
'
'
'
' Erforderliche Dateien/Verkn�pfungen:
' ------------------------------------
' keine
'
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
' # IsLeapYear()                �berpr�ft, ob es sich bei einem Jahr um ein Schaltjahr handelt
'   Eingabewerte:
'       year (erforderlich)     Jahr als Integer
'   R�ckgabewerte:
'       IsLeapYear              Boolean
'
'
' # DateWithoutSeparators()     Erm�glicht die Umwandlung von Datumseingaben im Format TTMMJJJJ
'                               ohne Trennzeichen
'   Eingabewerte:
'       rawdate (erforderlich)  Datumseingabe ohne Trennzeichen als String
'       separator (optional)    Trennzeichen als String
'   R�ckgabewerte:
'       DateWithoutSerials      Date
'   Fehlernummern:
'       10                      String ist zu kurz oder zu lang (muss 8 Zeichen haben)
'       11                      Monat ung�ltig (gr��er als 12)
'       12                      Tag ung�ltig (gr��er als 31)
'       14                      Mehr Tage angegeben, als Monat lang ist.
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Function IsLeapYear(year As Integer) As Boolean
' Anders als Excel ist 1900 nach dieser Funktion kein Schaltjahr!

If (year Mod 4 = 0 And year Mod 100 <> 0) Or (year Mod 400 = 0) Then
    IsLeapYear = True
Else
    IsLeapYear = False
End If
 
End Function

Public Function DateWithoutSeparators(rawdate As String, Optional separator As String) As Date

Dim intTag As Integer
Dim intMonat As Integer
Dim intJahr As Integer

' Falls das Datum bereits korrekt eingegeben wurde (Format TT.MM.JJJJ),
' muss keine Aufbereitung erfolgen, also Programmabbruch

If IsDate(rawdate) Then
    DateWithoutSeparators = rawdate
    Exit Function
End If

' Auslesen der Benutzereingabe und aufteilen

intTag = CInt(Left(rawdate, 2))
intMonat = CInt(Mid(rawdate, 3, 2))
intJahr = CInt(Right(rawdate, 4))

' Abfangen offensichtlich falscher Eingaben:

' - Eingabe hat weniger als 8 Zeichen, Datum ist unvollst�ndig
If Len(rawdate) <> 8 Then
    Err.Raise Number:=10, _
        Description:="Kann Datum nicht erzeugen, Eingabe zu kurz."
    Exit Function
End If

' - ung�ltiger Monat, Eingabe falsch bzw. in falscher Reihenfolge
If intMonat > 12 Then
    Err.Raise Number:=11, _
        Description:="Monat ung�ltig (gr��er als 12)."
    Exit Function
End If

' - ung�ltiger Tag, Eingabe falsch bzw. in falscher Reihenfolge
If intTag > 31 Then
    Err.Raise Number:=12, _
        Description:="Tag ung�ltig (gr��er als 31)."
    Exit Function
End If

' - ung�ltiger Tag (Monatsl�nge �berschritten), Eingabe falsch
Select Case intMonat
    Case 1, 3, 5, 7, 8, 10, 12:
    ' Diese Monate haben maximal 31 Tage
        If intTag > 31 Then
            Err.Raise Number:=14, _
                Description:="Mehr Tage angegeben, als Monat lang ist."
            Exit Function
        End If
    Case 4, 6, 9, 11:
    ' Diese Monate haben maximal 30 Tage
        If intTag > 30 Then
            Err.Raise Number:=14, _
                Description:="Mehr Tage angegeben, als Monat lang ist."
            Exit Function
        End If
    Case 2:
    ' Der Februar hat in Schaltjahren 29 und in �brigen Jahren 28 Tage.
    ' Daher erst ermitteln, ob ein Schaltjahr vorliegt:
        If (intJahr Mod 4 = 0 And intJahr Mod 100 <> 0) Or (intJahr Mod 400 = 0) Then
            If intTag > 29 Then
                Err.Raise Number:=14, _
                    Description:="Mehr Tage angegeben, als Monat lang ist."
                Exit Function
            End If
        Else
            If intTag > 28 Then
                Err.Raise Number:=14, _
                    Description:="Mehr Tage angegeben, als Monat lang ist."
                Exit Function
            End If
        End If
End Select

' Wenn keine Fehler aufgetreten sind, Eingabe neu zusammensetzen

DateWithoutSeparators = DateSerial(intJahr, intMonat, intTag)

Exit Function

End Function


