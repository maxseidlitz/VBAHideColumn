'''did that by myself. EASY CODE TO HIDE WHILE U HAVE A STATIC COLUMN TO HIDE'''
Sub Spalte_verbergen()
    Columns("D").Hidden = True
    MsgBox ("Spalte verborgen.")
End Sub

Sub Spalte_anzeigen()
Columns("D").Hidden = False
MsgBox ("Spalte wird wieder angezeigt.")
End Sub

''' not my work, got that from http://codevba.com/excel/hide_column.htm#.XljQbG5Fwy8 '''
Sub Column_Hide_Macro(control As IRibbonControl, ByRef CancelDefault)
     Columns("D").Hidden = True
    MsgBox ("You have hidden a column")

    ' You may put your code here
    ' to check if your monitored row is hidden

    CancelDefault = False   ' This enables the default action to continue
End Sub

Sub Column_UnHide_Macro(control As IRibbonControl, ByRef CancelDefault)
     Columns("D").Hidden = False
    MsgBox ("You have unhidden a column")

    ' You may put your code here
    ' to check if your monitored row is unhidden

    CancelDefault = False   ' This enables the default action to continue
End Sub

