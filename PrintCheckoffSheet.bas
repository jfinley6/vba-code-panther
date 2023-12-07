Attribute VB_Name = "PrintCheckoffSheet"
Sub Print_Screen()

    Dim sh As Worksheet
    Dim quantity As Integer
    
    For Each sh In ThisWorkbook.Worksheets
    
        If sh.Name = "Sheet1" Then
        quantity = sh.Range("BJ3")
        For i = 1 To quantity
            sh.PrintOut
        Next i
        
        End If
    Next sh

End Sub
