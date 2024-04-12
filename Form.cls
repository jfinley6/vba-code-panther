VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim shModel      As Worksheet
    Dim sOrderNumber As String
    Set shForm = ThisWorkbook.Sheets("Form")
    If Target.Address = "$G$5" Then
        sOrderNumber = Range("G5")
        If sOrderNumber <> "" Then
            Call Duplicate_Search(sOrderNumber)
        End If
    End If
End Sub
