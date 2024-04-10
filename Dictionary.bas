Attribute VB_Name = "Dictionary"
Sub Firmware_Dictionary(sPantherModel)

Dim MyTable As ListObject
Dim myArray As Variant
Dim i As Long
Dim x As Long
Dim lColumnCount As Long
Dim bFirmwareFound As Boolean
Dim sCurrentFirmware As String

Dim testNum As Integer:
testNum = 0
bFirmwareFound = False

Dim dict: Set dict = CreateObject("Scripting.Dictionary")

'Check for duplicate model names
Find_Duplicate_Values_From_Dictionary

'Set path for Table variable
Set MyTable = Sheets("FIRMWARE DICTIONARY").ListObjects("FIRMWARE_DICTIONARY")
  
'Set number of Table columns
lColumnCount = MyTable.DataBodyRange.Columns.Count

'Create Array List from Table
myArray = MyTable.DataBodyRange

'Loop through every item in each column and see if matching firmware exists
'TODO: Add Logic to Create Dictionary from Table
For i = 1 To lColumnCount
    For x = LBound(myArray) To UBound(myArray)
        sCurrentFirmware = MyTable.ListColumns(i).Name
        'Check To See if Cell is Empty
        If Not Trim(myArray(x, i) & vbNullString) = vbNullString Then
            If myArray(x, i) = sPantherModel Then bFirmwareFound = True: sModelName = sCurrentFirmware: Exit For
        End If
        Next x
    If bFirmwareFound Then firmwareExists = True: Exit For
    Next i
    If Not bFirmwareFound Then firmwareExists = False
    
    
End Sub

Sub test()
    For N = 1 To Range("FIRMWARE_DICTIONARY[]").Cells(i, 1)
        Debug.Print
    Next N

End Sub

Sub Open_Dictionary()

    ThisWorkbook.Worksheets("FIRMWARE DICTIONARY").Activate

End Sub

Sub Find_Duplicate_Values_From_Dictionary()

Dim myRange As Range
Dim i As Integer
Dim j As Integer
Dim myCell As Range
Dim iOriginalCellColor As Integer
Dim sDuplicateCells As String
Dim duplicateFound As Boolean

Set myRange = Range("FIRMWARE_DICTIONARY")
duplicateFound = False

For Each myCell In myRange
    If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
        duplicateFound = True
        iOriginalCellColor = myCell.Interior.ColorIndex
        myCell.Interior.ColorIndex = 3
        sDuplicateCells = sDuplicateCells + Replace(myCell.Address, "$", "") + " "
    End If
Next

If duplicateFound Then

    ThisWorkbook.Worksheets("FIRMWARE DICTIONARY").Activate
    
    shForm.Range("I5") = ""

    MsgBox "Duplicates can be found at the following cells: " & sDuplicateCells & vbCrLf & _
        "Please remove duplicates and try again.", vbCritical
    For Each myCell In myRange
        If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
           myCell.Interior.ColorIndex = iOriginalCellColor
        End If
    Next
    End
End If

End Sub

Sub Add_To_Dictionary(columnName, newValue As String)

Dim ws As Worksheet
Dim lastRow As Long
Dim columnRange As Range
Dim targetCell As Range
Dim headerRow As Range
Dim headerCell As Range
Dim i As Long

Set ws = ThisWorkbook.Sheets("FIRMWARE DICTIONARY")

Set headerRow = ws.Rows(1)
Set headerCell = headerRow.Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole)

If Not headerCell Is Nothing Then
    Set columnRange = ws.Columns(headerCell.Column)
    
    lastRow = ws.Cells(ws.Rows.Count, headerCell.Column).End(xlUp).Row
    
    For i = lastRow To 1 Step -1
        If Not IsEmpty(columnRange.Cells(i, 1)) Then
            Set targetCell = columnRange.Cells(i + 1, 1)
            Exit For
        End If
    Next i
    
    targetCell.Value = newValue

End If

End Sub

