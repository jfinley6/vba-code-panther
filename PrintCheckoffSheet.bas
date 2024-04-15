Attribute VB_Name = "PrintCheckoffSheet"
'@IgnoreModule ModuleWithoutFolder
Option Explicit 'Force explicit variable declaration.
Public Sub Print_Screen()

    Dim sheet As Worksheet

    For Each sheet In ThisWorkbook.Worksheets
    
        If sheet.Name = "CHECK SHEET" Then
            sheet.PrintOut
        End If
    Next sheet

End Sub
