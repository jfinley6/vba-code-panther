Attribute VB_Name = "Dictionary"
Sub Firmware_Dictionary(sPantherModel)

    Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer
    Dim sSplitString() As String
    Dim sModelValue As String
    
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    
    myFile = "H:\Service Department\Production Forms\firmware.txt"
    
    Open myFile For Input As #1
    
     Do Until EOF(1)
        Line Input #1, textline
        If InStr(textline, "~") <> 0 Then
            sModelValue = Replace(Trim(textline), "~", "")
        ElseIf InStr(textline, "/") = 0 And CStr(textline) <> "" And InStr(textline, "~") = 0 Then
            If dict.Exists(Trim(textline)) Then
                MsgBox (Trim(textline) & " is Duplicated in the Dictionary. Please Fix and Try Again")
                shForm.Range("I5") = ""
                End
            Else
                dict.Add Trim(textline), Trim(sModelValue)
            End If
        End If
        
    Loop
    Close #1
    
    If dict.Exists(sPantherModel) Then
        firmwareExists = True
        sModelName = dict.Item(sPantherModel)
    ElseIf InStr(sPantherModel, "STAND") <> 0 Then
        firmwareExists = True
        sModelName = "STAND"
    ElseIf InStr(sPantherModel, "MNS") <> 0 Then
        firmwareExists = True
        sModelName = "MNS"
    Else
        firmwareExists = False
    End If
    
End Sub

Sub Open_Dictionary()

    Dim fso As Object
    Dim sfile As String
    Set fso = CreateObject("shell.application")
    sfile = "H:\Service Department\Production Forms\firmware.txt"
    fso.Open (sfile)

End Sub

