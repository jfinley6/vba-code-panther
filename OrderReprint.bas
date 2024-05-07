Attribute VB_Name = "OrderReprint"
Sub Repopulate_Form(Target, Cancel)
    Dim WorkSheetRange As String
    WorkSheetRange = OrderReprint.Get_Active_Worksheet_Range
    
    'Check to See If User Right Clicked in the Table
    If Not Intersect(Target, Range(WorkSheetRange)) Is Nothing Then '<- Check column is in table
      With Intersect(Target.EntireRow, ActiveSheet.UsedRange)
        'Check if Order is Blank
        
        If Target.EntireRow.Cells(1, 3) <> "" And Target.Column = "1" Then
            Call Copy_Table_Values(Target)
            Cancel = True
            Form.Activate
        End If
      End With
    End If

End Sub

'Return the Column Range of the Clicked Table
Function Get_Active_Worksheet_Range() As String
    Dim WorkSheetName As String
    WorkSheetName = ActiveSheet.Name
    
    Select Case WorkSheetName
    Case "P9", "P5c", "FLEX", "STAND"
        Get_Active_Worksheet_Range = "A:J"
    Case "SHADOW", "MNS"
        Get_Active_Worksheet_Range = "A:I"
    End Select
End Function

Sub Copy_Table_Values(Target)
    Dim TargetSheet As Worksheet
    Set TargetSheet = ThisWorkbook.Sheets("Form")
    Dim WorkSheetName As String
    WorkSheetName = ActiveSheet.Name
    
    'Reset Form Values
    TargetSheet.Range("G5, G7, G9, G12:I17").Value = ""
    
    'Assign Values to Form Cell
    TargetSheet.Range("G5") = Target.EntireRow.Cells(1, 3).Value 'Order Number
    TargetSheet.Range("G7") = Target.EntireRow.Cells(1, 4).Value 'Customer Name
    TargetSheet.Range("G9") = Target.EntireRow.Cells(1, 5).Value 'End User
    
    TargetSheet.Range("G12") = Target.EntireRow.Cells(1, 6).Value 'Model
    TargetSheet.Range("H12") = "1" 'Quantity
    
    'STAND Doesn't Have Label Size So Don't Copy Value
    If WorkSheetName <> "STAND" Then
        TargetSheet.Range("I12") = Target.EntireRow.Cells(1, 7).Value
    End If

End Sub
