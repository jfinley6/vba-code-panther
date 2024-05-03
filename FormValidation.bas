Attribute VB_Name = "FormValidation"
Sub Duplicate_Search(sOrderNumber)
    
    Dim rgFound As Range
    Dim iConfirmDuplicate
    
    Set rgFound = ThisWorkbook.Sheets("P9").Range("C:C").Find(sOrderNumber)
    If rgFound Is Nothing Then
        Set rgFound = ThisWorkbook.Sheets("P5c").Range("C:C").Find(sOrderNumber)
        If rgFound Is Nothing Then
            Set rgFound = ThisWorkbook.Sheets("FLEX").Range("C:C").Find(sOrderNumber)
            If rgFound Is Nothing Then
                Set rgFound = ThisWorkbook.Sheets("SHADOW").Range("C:C").Find(sOrderNumber)
                If rgFound Is Nothing Then
                    Set rgFound = ThisWorkbook.Sheets("STAND").Range("C:C").Find(sOrderNumber)
                    If rgFound Is Nothing Then
                        Set rgFound = ThisWorkbook.Sheets("MNS").Range("C:C").Find(sOrderNumber)
                            If rgFound Is Nothing Then
                            
                            Else
                                iConfirmDuplicate = MsgBox("Order Number already exists in MNS : " & rgFound.Address & vbNewLine & vbNewLine _
                                & "Submitting will create duplicate serial numbers For this order" & vbNewLine _
                                & "Are you sure you want To proceed?", vbYesNo + vbQuestion, "WARNING: Duplicate Order Found")
                                If iConfirmDuplicate = vbNo Then
                                    shForm.Range("I5") = ""
                                    Exit Sub
                                End If
                            End If
                    Else
                        iConfirmDuplicate = MsgBox("Order Number already exists in STAND : " & rgFound.Address & vbNewLine & vbNewLine _
                        & "Submitting will create duplicate serial numbers For this order" & vbNewLine _
                        & "Are you sure you want To proceed?", vbYesNo + vbQuestion, "WARNING: Duplicate Order Found")
                        If iConfirmDuplicate = vbNo Then
                            shForm.Range("I5") = ""
                            Exit Sub
                        End If
                    End If
                Else
                    iConfirmDuplicate = MsgBox("Order Number already exists in SHADOW : " & rgFound.Address & vbNewLine & vbNewLine _
                    & "Submitting will create duplicate serial numbers For this order" & vbNewLine _
                    & "Are you sure you want To proceed?", vbYesNo + vbQuestion, "WARNING: Duplicate Order Found")
                    If iConfirmDuplicate = vbNo Then
                        shForm.Range("I5") = ""
                        Exit Sub
                    End If
                End If
            Else
                iConfirmDuplicate = MsgBox("Order Number already exists in FLEX : " & rgFound.Address & vbNewLine & vbNewLine _
                & "Submitting will create duplicate serial numbers For this order" & vbNewLine _
                & "Are you sure you want To proceed?", vbYesNo + vbQuestion, "WARNING: Duplicate Order Found")
                If iConfirmDuplicate = vbNo Then
                    shForm.Range("I5") = ""
                    Exit Sub
                End If
            End If
        Else
            iConfirmDuplicate = MsgBox("Order Number already exists in P5c : " & rgFound.Address & vbNewLine & vbNewLine _
            & "Submitting will create duplicate serial numbers For this order" & vbNewLine _
            & "Are you sure you want To proceed?", vbYesNo + vbQuestion, "WARNING: Duplicate Order Found")
            If iConfirmDuplicate = vbNo Then
                shForm.Range("I5") = ""
                Exit Sub
            End If
        End If
    Else
        iConfirmDuplicate = MsgBox("Order Number already exists in P9 : " & rgFound.Address & vbNewLine & vbNewLine _
        & "Submitting will create duplicate serial numbers For this order" _
        & vbNewLine & "Are you sure you want To proceed?", vbYesNo + vbQuestion, "WARNING: Duplicate Order Found")
        If iConfirmDuplicate = vbNo Then
            shForm.Range("I5") = ""
            Exit Sub
        End If
    End If
        
End Sub

