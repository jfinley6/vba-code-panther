Attribute VB_Name = "PrintCheckoffSheet"
Sub Print_Screen()

    Dim sh As Worksheet
    Dim quantity As Integer

    For Each sh In ThisWorkbook.Worksheets
    
        If sh.Name = "CHECK SHEET" Then
            sh.PrintOut
        End If
    Next sh

End Sub
