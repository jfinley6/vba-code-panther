VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addNewModelForm 
   Caption         =   "Add Model To Dictionary"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5220
   OleObjectBlob   =   "addNewModelForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addNewModelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable declarations necessary to place image in title
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As Any) As LongPtr
Private Const WM_SETICON As Long = &H80
Private Const ICON_SMALL As LongPtr = 0&
Private Const ICON_BIG As LongPtr = 1&

Private Sub UserForm_Initialize()

    Dim ModelNumber As String
    
    ModelNumber = "PA-P9-118e"
    
    With ModelListBox
       .AddItem "Predator Straight Tamp (WAGO)"
       .AddItem "Predator Straight Tamp (Beijer)"
       .AddItem "Predator Swing Arm"
       .AddItem "Phantom"
       .AddItem "P5c"
       .AddItem "Flex"
       .AddItem "Shadow"
       .ListIndex = 0 'Sets the default value
    End With
    
    InstructionLabel1.Caption = "The model number below does not exist in the dictionary:"
    
    With ModelNumberLabel
        .Caption = ModelNumber
        .FontBold = True
    End With
    
    InstructionLabel2.Caption = "Please select the corresponding machine from the list and click add"
    
    Call SetIconFromImageControl
    
End Sub

Private Sub CancelButton_Click()
    Unload Me
    End
End Sub

Private Sub SetIconFromImageControl()
    On Error GoTo errExit
    Dim hWnd As LongPtr, hIcon As LongPtr
    hWnd = FindWindow("ThunderDFrame", Caption)
    hIcon = ImageForIcon.Picture.Handle
    If hWnd <> 0 And hIcon <> 0 Then
        SendMessage hWnd, WM_SETICON, ICON_SMALL, hIcon
        SendMessage hWnd, WM_SETICON, ICON_BIG, hIcon
    End If
errExit:
End Sub
