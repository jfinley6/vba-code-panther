Attribute VB_Name = "PrintCheckoffSheet"
Sub Print_Screen()

    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name = "Sheet1" Then
            sh.PrintOut
        End If
    Next sh

End Sub
