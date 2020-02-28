Sub Spalte_verbergen_ORIGINAL()
    '''Deklaration
    Dim varSpalte As String
    Dim varDatum As String
    
    '''Initialisierung
    Let varDatum = Range("A1").Value
    
        'Let varSpalte = Val(Range("C2").Value) 'nimmt den Wert der Spalzte C2 auf für Variable varSpalte
        'MsgBox (varDatum) ''Test der Variable varSpalte
        
    '!!!T E S T!!! variable Datum
    MsgBox (varDatum)
    
    '''je nach Datum (Monat) werden die irrelevanten Spalten ausgeblendet
    Select Case varDatum
        Case 1, 1 To 31, 1
          Columns("B").EntireColumn.Hidden = True '"False" zeigt Spalte wieder an
          Columns("E:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
        
        Case 1, 2 To 31, 2
          Columns("B:D").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("G:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 1, 3 To 31, 3
          Columns("B:F").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("I:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 1, 4 To 31, 4
          Columns("B:H").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("K:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 1, 5 To 31, 5
          Columns("B:J").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("M:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 1, 6 To 31, 6
          Columns("B:L").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("O:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 1, 7 To 31, 7
          Columns("B:N").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("Q:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 1, 8 To 31, 8
          Columns("B:P").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("S:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 1, 9 To 31, 9
          Columns("B:R").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("U:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
          
        Case 1, 10 To 31, 10
          Columns("B:T").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("W:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")
        
        Case 1, 11 To 31, 11
          Columns("B:V").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          Columns("Y:Z").EntireColumn.Hidden = True
          MsgBox ("Spalten verborgen.")

        Case 1, 12 To 31, 12
          Columns("B:X").EntireColumn.Hidden = True  '"False" zeigt Spalte wieder an
          MsgBox ("Spalten verborgen.")
    End Select

End Sub

Sub Spalten_alleEinblenden()
    '''Deklaration
    Dim varSpalte As String
    Dim varDatum As String
    
    '''Initialisierung
    Let varDatum = Val(Range("A1").Value)
    
    '!'!'! TESTOBJEKT!'!'!'
    '''je nach Datum (Monat) werden die irrelevanten Spalten ausgeblendet
   
          Range("B:Z").EntireColumn.Hidden = False '"False" sorgt für Wiedererscheinen
          MsgBox ("Alle Spalten angezeigt.")
    
    
End Sub

Sub Spalte_anzeigen()
    Dim varSpalte As String
     Let varSpalte = Column.Name
        Columns("D").Hidden = False
        MsgBox ("Spalte wird wieder angezeigt.")
End Sub
