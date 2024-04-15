Attribute VB_Name = "Dictionary"
'@IgnoreModule ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess, ProcedureNotUsed, ModuleWithoutFolder
Option Explicit 'Force explicit variable declaration.

Public Sub Firmware_Dictionary(ByVal sPantherModel As String)

Dim MyTable As ListObject
Dim myArray As Variant
Dim i As Long
Dim X As Long
Dim lColumnCount As Long
Dim bFirmwareFound As Boolean
Dim sCurrentFirmware As String

bFirmwareFound = False

'Check for duplicate model names
Find_Duplicate_Values_From_Dictionary

'Set path for Table variable
Set MyTable = FirmwareDictionary.ListObjects("FIRMWARE_DICTIONARY")
  
'Set number of Table columns
lColumnCount = MyTable.DataBodyRange.Columns.Count

'Create Array List from Table
myArray = MyTable.DataBodyRange

'Loop through every item in each column and see if matching firmware exists
'TODO: Add Logic to Create Dictionary from Table
For i = 1 To lColumnCount
    For X = LBound(myArray) To UBound(myArray)
        sCurrentFirmware = MyTable.ListColumns(i).Name
        'Check To See if Cell is Empty
        If Not Trim$(myArray(X, i) & vbNullString) = vbNullString Then
            If myArray(X, i) = sPantherModel Then bFirmwareFound = True: sModelName = sCurrentFirmware: Exit For
        End If
        Next X
    If bFirmwareFound Then firmwareExists = True: Exit For
    Next i
    If Not bFirmwareFound Then firmwareExists = False
    
    
End Sub

Public Sub Open_Dictionary()

    FirmwareDictionary.Activate

End Sub

Public Sub Find_Duplicate_Values_From_Dictionary()

Dim myRange As Range
Dim myCell As Range
Dim iOriginalCellColor As Long
Dim sDuplicateCells As String
Dim duplicateFound As Boolean

Set myRange = FirmwareDictionary.Range("FIRMWARE_DICTIONARY")
duplicateFound = False

For Each myCell In myRange
    If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
        duplicateFound = True
        iOriginalCellColor = myCell.Interior.ColorIndex
        myCell.Interior.ColorIndex = 3
        sDuplicateCells = sDuplicateCells + Replace(myCell.Address, "$", vbNullString) + " "
    End If
Next

If duplicateFound Then

    FirmwareDictionary.Activate
    
    shForm.Range("I5") = vbNullString

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

Public Sub Add_To_Dictionary(ByVal columnName As String, ByVal newValue As String)

Dim ws As Worksheet
Dim lastRow As Long
Dim columnRange As Range
Dim targetCell As Range
Dim headerRow As Range
Dim headerCell As Range
Dim i As Long

Set ws = FirmwareDictionary

Set headerRow = ws.Rows(1)
Set headerCell = headerRow.Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole)

If Not headerCell Is Nothing Then
    Set columnRange = ws.Columns(headerCell.Column)
    
    'Go To The Bottom Of The Column
    lastRow = ws.Cells(ws.Rows.Count, headerCell.Column).End(xlUp).Row
    
    'Find First Empty Cell Where New Value Should Go
    For i = lastRow To 1 Step -1
        If Not IsEmpty(columnRange.Cells(i, 1)) Then
            Set targetCell = columnRange.Cells(i + 1, 1)
            Exit For
        End If
    Next i
    
    'Set Value Of Target Cell
    targetCell.Value = newValue
    
    'Alert User Pairing Was Successful
    If targetCell.Value = newValue Then
        MsgBox newValue & " has been paired with " & headerCell & " successfully!"
    End If
    

End If

End Sub


