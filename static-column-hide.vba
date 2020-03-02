Sub TagesdatumTest01()
'Deklaration
Dim nowTagesdatum As String
Let nowTagesdatum = Now                 'hier wird in der Variable <nowTagesdatum> das aktuelle Datum + Uhrzeit im Format TTMMJJ MMHHSS gespeichert

    MsgBox Month(nowTagesdatum)         'Ausgabe Datum

End Sub

Sub Spalte_verbergen_ORIGINAL()
    '''Deklaration
    Dim strDatum As String
    Dim bytDatum As Byte
    
    '''Initialisierung
    Let strDatum = Now                  'hier wird auf die Variable <nowTagesdatum> das aktuelle Datum + Uhrzeit im Format TTMMJJ MMHHSS initialisiert
    
    Let strDatum = Month(strDatum)      'String wird nur auf Monat minimiert
    Let bytDatum = strDatum             'String wird von Zeichenkette ("Wort") zu definierter Zahl '!' mit Typumwandlung hat es nicht funktionert '!'
    Let bytDatum = bytDatum - 1         'aktueller Monat wird - 1 gerechnet um den Monat der Inventur anzuzeigen; Logik: <Inventurmonat = aktueller Monat - 1>
    
        '!!!T E S T!!! variable Datum; ÜBERPRÜFUNG DER INITIALISIERUNG
        MsgBox ("Sie befinden sich nun im Monat " & bytDatum & ".")
        
    '''je nach Datum (Monat) werden die irrelevanten Spalten ausgeblendet
    Select Case bytDatum
        Case 1
          Columns("B").EntireColumn.Hidden = True '"= <False>" zeigt Spalte wieder an
          Columns("E:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
        
        Case 2
          Columns("B:D").EntireColumn.Hidden = True
          Columns("G:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 3
          Columns("B:F").EntireColumn.Hidden = True
          Columns("I:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 4
          Columns("B:H").EntireColumn.Hidden = True
          Columns("K:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 5
          Columns("B:J").EntireColumn.Hidden = True
          Columns("M:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 6
          Columns("B:L").EntireColumn.Hidden = True
          Columns("O:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 7
          Columns("B:N").EntireColumn.Hidden = True
          Columns("Q:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 8
          Columns("B:P").EntireColumn.Hidden = True
          Columns("S:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 9
          Columns("B:R").EntireColumn.Hidden = True
          Columns("U:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 10
          Columns("B:T").EntireColumn.Hidden = True
          Columns("W:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
        
        Case 11
          Columns("B:V").EntireColumn.Hidden = True
          Columns("Y:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")

        Case 12
          Columns("B:X").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
    End Select

End Sub

Sub Spalten_alleEinblenden()
    '''Deklaration
    Dim varSpalte As String
    Dim varDatum As String
    
       
        Range("B:Z").EntireColumn.Hidden = False '"False" sorgt für Wiedererscheinen
        MsgBox ("Alle Spalten angezeigt.")
    
    
End Sub
