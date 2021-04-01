Attribute VB_Name = "Module11"
Sub Max_offsett()
 
    Dim mtch As range
    Dim ticker As String
    Dim volume As String
    Dim range As range
    Dim High_Percent As Double
    Dim Percent_Change As range
    Set range = Application.range("K:K")
    High_Percent = Application.WorksheetFunction.max(range)
    mtch = High_Percent
    
    For Each mtch In range
            mtch = Percent_Change.Value
        If mtch = High_Percent Then
            ticker = Percent_Change.Offset(, -2).Value
            range("O" & 2).Value = ticker
            volume = Percent_Change.Offset(, 1).Value
            range("P" & 2).Value = volume
    Exit For

         End If
         
    Next mtch
 
End Sub
