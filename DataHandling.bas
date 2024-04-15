Attribute VB_Name = "DataHandling"
'Variables used in multiple subroutines
Public shForm       As Worksheet
Public sOrderNumber As String
Public sCustomer    As String
Public sEndUser     As String
Public sPantherModel As String
Public sLabelSize   As String
Public sPrinterName As String
Public sPrinterIP   As String
Public sSerialNumber As String
Public sPantherPLC  As String
Public sPantherOptions As String
Public sCustomFirmware As String

Public sLabelPart1  As String
Public sLabelPart2  As String
Public sLabelPart3  As String
Public sLabelPart4  As String
Public sLabelPart5  As String
Public sLabelPart6  As String
Public sLabelPart7  As String

Public sLabelZPL    As String

Public iMachinesOrdered As Integer
Public iCurrent     As Integer
Public iNumOrdered  As Integer
Public iMachineRow  As Integer

Public firmwareExists As Boolean
Public sModelName As String

Public sPLCFirmware As String
Public sPLCFileExt As String
Public sHMIFileExt As String
Public sHMIFirmware As String
Public sSystemFirmware As String
Public sSystemExt  As String
Public sServoFirmware As String
Public sServoExt   As String
Public bSlideOptions As Boolean

Public sSelectedMachineName As String

'Examine Tag Checkboxes to see which labels to print
Sub Get_Tag_Checkbox()
    If shForm.CheckBoxes("Serial Number Checkbox").Value = 1 And shForm.CheckBoxes("Toe Tag Checkbox").Value = 1 Then
            Print_Machine_Tags sLabelZPL, sPrinterIP
            If InStr(sPantherModel, "SLIDE") = 0 Then
                Call Get_Serial_Number_Label(sCustomer, sPantherModel, sPantherPLC, sSerialNumber, iMachinesOrdered, sPrinterIP, iCurrent)
            End If
        ElseIf shForm.CheckBoxes("Serial Number Checkbox").Value = -4146 And shForm.CheckBoxes("Toe Tag Checkbox").Value = 1 Then
            Call Print_Machine_Tags(sLabelZPL, sPrinterIP)
        ElseIf shForm.CheckBoxes("Serial Number Checkbox").Value = 1 And shForm.CheckBoxes("Toe Tag Checkbox").Value = -4146 Then
            If InStr(sPantherModel, "SLIDE") = 0 Then
                Call Get_Serial_Number_Label(sCustomer, sPantherModel, sPantherPLC, sSerialNumber, iMachinesOrdered, sPrinterIP, iCurrent)
            End If
    End If
        
End Sub

'Get Strings for Toe Tag ZPL
Sub Get_Label_Strings(sPrinterName)
    Select Case sPrinterName
        Case "Service"
            sPrinterIP = "192.168.17.97"
            sLabelPart1 = "^FS^FB500,2^CF0,80,80^FT40,85^A0N,30,40^FD"
            sLabelPart2 = "^FS^FB600,2^CF0,80,80^FT40,155^A0N,30,40^FD"
            sLabelPart3 = "^FS^FT40,200^A0N,30,40^FD"
            sLabelPart4 = "^FS^FT40,250^A0N,30,40^FD"
            sLabelPart5 = "^FS^FT40,300^A0N,30,40^FD"
            sLabelPart6 = "^FS^FT300,300^A0N,30,40^FD"
            sLabelPart7 = "^FS^FT300,250^A0N,30,40^FD"
        Case "Darkside"
            sPrinterIP = "192.168.17.33"
            sLabelPart1 = "^FS^FB800,2^FT60,100^A0N,44,59^FD"
            sLabelPart2 = "^FS^FB800,2^FT60,210^A0N,44,59^FD"
            sLabelPart3 = "^FS^FT60,280^A0N,44,59^FD"
            sLabelPart4 = "^FS^FT60,350^A0N,44,59^FD"
            sLabelPart5 = "^FS^FT60,420^A0N,44,59^FD"
            sLabelPart6 = "^FS^FT520,420^A0N,44,59^FD"
            sLabelPart7 = "^FS^FT520,350^A0N,44,59^FD"
    End Select
End Sub

'Print out Serial Number Label Depending on Form Inputs
Sub Get_Serial_Number_Label(sCustomerName, sPantherModel, sPantherPLC, sSerialNumber, iMachinesOrdered, sPrinterIP, iCurrent)
    Dim sOptionAHC  As String
    Dim sOptionEXP  As String
    Dim sSerialNumberZPL As String
    Dim sCurrentIP  As String
    
    sCurrentIP = sPrinterIP
    sCustomFirmware = shForm.Range("D15")
    
    'Reset Values
    sPLCFirmware = ""
    sPLCFileExt = ""
    sHMIFirmware = ""
    sHMIFileExt = ""
    sServoFirmware = ""
    sServoExt = ""
    sSystemFirmware = ""
    sSystemExt = ""
    
    Select Case sModelName
    Case "Predator Straight Tamp (WAGO)"
        sPLCFirmware = shForm.Range("O6")
        sPLCFileExt = shForm.Range("Q6")
        sHMIFirmware = shForm.Range("O7")
        sHMIFileExt = shForm.Range("Q7")
        sServoFirmware = "SERVO: " + shForm.Range("O8")
    Case "Predator Straight Tamp (Beijer)"
        sPLCFirmware = shForm.Range("T6")
        sHMIFirmware = shForm.Range("T7")
        sHMIFileExt = shForm.Range("V7")
        sServoFirmware = "SERVO: " + shForm.Range("T8")
    Case "Predator Swing Arm"
        sPLCFirmware = shForm.Range("O11")
        sPLCFileExt = shForm.Range("Q11")
        sHMIFirmware = shForm.Range("O12")
        sHMIFileExt = shForm.Range("Q12")
        sServoFirmware = "SERVO: " + shForm.Range("O13")
    Case "Phantom"
        sPLCFirmware = shForm.Range("T11")
        sPLCFileExt = shForm.Range("V11")
        sHMIFirmware = shForm.Range("T12")
        sHMIFileExt = shForm.Range("V12")
    Case "Flex"
        sPLCFirmware = shForm.Range("O16")
        sPLCFileExt = shForm.Range("Q16")
        sHMIFirmware = shForm.Range("O17")
        sHMIFileExt = shForm.Range("Q17")
    Case "P5c"
        sPLCFirmware = shForm.Range("T16")
        sPLCFileExt = shForm.Range("V16")
        sHMIFirmware = shForm.Range("T17")
        sHMIFileExt = shForm.Range("V17")
    Case "Shadow"
        sServoFirmware = shForm.Range("O21")
        sServoExt = shForm.Range("Q21")
    Case "STAND"
        sSystemFirmware = shForm.Range("T21")
        sSystemExt = shForm.Range("V21")
    Case "MNS"
        sServoFirmware = "9-12-17"
    End Select
    
    Select Case sPrinterName
        Case "Service"
            sLabelPart1 = "^XA^FO50,180^GB544,0,2^FS^FO50,230^GB544,0,2^FS^FO50,280^GB544,0,2^FS"
            sLabelPart2 = "^FO10,55^CF0,80,80^FB600,1,0,C^FD"
            sLabelPart3 = "^FS^CF0,30,30^FO80,146^FB650,1,,L,^FD"
            sLabelPart4 = "^FS^CF0,30,30^FO80,196^FB650,1,,L, ^FD"
            sLabelPart5 = "^FS^CF0,30,30^FO80,246^FB650,1,,L, ^FD"
        Case "Darkside"
            sLabelPart1 = "^XA^FO90,235^GB800,2,2^FS^FO90,310^GB800,2,2^FS^FO90,385^GB800,2,2^FS"
            sLabelPart2 = "^FO20,35^CF0,118,118^FB900,1,0,C^FD"
            sLabelPart3 = "^FS^CF0,44,44^FO100,190^FB800,1,,L,^FD"
            sLabelPart4 = "^FS^CF0,44,44^FO100,265^FB800,1,,L, ^FD"
            sLabelPart5 = "^FS^CF0,44,44^FO100,340^FB650,1,,L, ^FD"
    End Select
    
    'Check sPantherModel to Determine Structure of the Serial Number Label
    If sModelName = "Shadow" Then
        sSerialNumberZPL = sLabelPart1 + sLabelPart2 + sSerialNumber + sLabelPart3 + "SERVO: " + sServoFirmware + sServoExt + sLabelPart4 + sLabelPart5 + "^FS^PQ2,0,1,Y,^XZ"
    ElseIf sModelName = "STAND" Then
        sSerialNumberZPL = sLabelPart1 + sLabelPart2 + sSerialNumber + sLabelPart3 + "SYSTEM: " + sSystemFirmware + sSystemExt + sLabelPart4 + sLabelPart5 _
        + "^FS^PQ2,0,1,Y^XZ"
    ElseIf InStr(sPantherModel, "MNS") <> 0 Then
        sSerialNumberZPL = sLabelPart1 + sLabelPart2 + sSerialNumber + sLabelPart3 + "SERVO: " + sServoFirmware + sLabelPart4 + sLabelPart5 _
        + "^FS^PQ2,0,1,Y^XZ"
    Else
        sSerialNumberZPL = sLabelPart1 + sLabelPart2 + sSerialNumber + sLabelPart3 + "PLC: " + sPLCFirmware + sPantherOptions + sOptionEXP _
        + IIf(Len(sCustomFirmware) > 0 And shForm.CheckBoxes("PLC Checkbox").Value = 1, sCustomFirmware, "") + sPLCFileExt + sLabelPart4 + "HMI: " _
        + sHMIFirmware + IIf(Len(sCustomFirmware) > 0 And shForm.CheckBoxes("HMI Checkbox").Value = 1, sCustomFirmware, "") _
        + sHMIFileExt + sLabelPart5 + sServoFirmware + sServoExt + "^FS^PQ2,0,1,Y^XZ"
    End If
    
    Call Print_Machine_Tags(sSerialNumberZPL, sCurrentIP)
    
End Sub

'Clear Form Page
Sub Reset_Form()
    
    Dim iMessage    As VbMsgBoxResult
    iMessage = MsgBox("Reset Form?", vbYesNo + vbQuestion, "Reset Confirmation")
    
    If iMessage = vbNo Then Exit Sub
    
    ThisWorkbook.Sheets("Form").Range("G7, G9, G12:I17").Value = ""
    
End Sub

Sub Submit_Form()
    
    'Variables
    Dim i               As Integer
    Dim sMacroInUse     As String
    Dim shModel         As Worksheet
    Dim loTable         As ListObject
    Dim NewRow          As ListRow
    Dim sTargetSheet    As String
    Dim aPantherModel() As String
    
    'Target sheets
    Set shForm = ThisWorkbook.Sheets("Form")
    
    sMacroInUse = shForm.Range("I5")
    
    If sMacroInUse = "Writing" Then
        MsgBox ("Serial Number Generator in use, please wait And try again")
        Exit Sub
    End If
    
    If shForm.CheckBoxes("Serial Number Checkbox").Value <> 1 Then
        MsgBox "Serial Number CheckBox Must Be Checked To Be Able To Submit Machines"
        End
    End If
    
    
    shForm.Range("I5") = "Writing"
    
    'Count number of rows with machines on order
    iMachinesOrdered = Application.WorksheetFunction.CountIf(shForm.Range("G12:G17"), "*")
    
    If iMachinesOrdered = 0 Then
        MsgBox ("There Are No Machines To Submit")
        shForm.Range("I5") = ""
        Exit Sub
    Else
        
        sOrderNumber = shForm.Range("G5")
        
        'Duplicate Order Search
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
        
        sCustomer = shForm.Range("G7")
        sEndUser = shForm.Range("G9")
        sPrinterName = shForm.Range("G19")
        
        For i = 1 To iMachinesOrdered
        
            iMachineRow = 11 + i
            
            sPantherModel = shForm.Range("G" & iMachineRow).Value
            Call Firmware_Dictionary(sPantherModel)
            If firmwareExists = False And InStr(sPantherModel, "SLIDE") = 0 Then
                addNewModelForm.Show
            End If
            sPantherOptions = shForm.Range("J" & iMachineRow).Value
            
            Validate_Options_Input
            
            iNumOrdered = shForm.Range("H" & iMachineRow)
            sLabelSize = shForm.Range("I" & iMachineRow)
            
            If Right(sPantherModel, 5) = "XP-LH" Or Right(sPantherModel, 2) = "XP" Or InStr(sPantherModel, "XP") <> 0 Then
                sPantherPLC = "XP"
            Else
                sPantherPLC = ""
            End If
            
            aPantherModel = Split(sPantherModel, "-")
            sTargetSheet = aPantherModel(1)
            
            If InStr(sPantherModel, "SLIDE") = 0 Then
                Set shModel = ThisWorkbook.Sheets(sTargetSheet)
            End If
            
            For iCurrent = 1 To iNumOrdered
            
                If InStr(sPantherModel, "SLIDE") = 0 Then
                    'Add new row and populate
                    Set loTable = shModel.ListObjects(sTargetSheet & "_table")
                    Set NewRow = loTable.ListRows.Add
                    
                    'Print Worksheet to Default Printer
                    Print_Screen
                    
                    With NewRow
                        'increment serial number
                        
                        If sTargetSheet = "P9" And Left(sPantherPLC, 2) = "XP" Then
                            .Range(1) = Application.WorksheetFunction.MaxIfs(Sheet2.UsedRange.Columns(1), Sheet2.UsedRange.Columns(1), ">=" & 50000) + 1
                        ElseIf sTargetSheet = "P9" And sPantherPLC <> "XP" Then
                            .Range(1) = Application.WorksheetFunction.MaxIfs(Sheet2.UsedRange.Columns(1), Sheet2.UsedRange.Columns(1), "<" & 50000) + 1
                        ElseIf InStr(sPantherModel, "STAND") <> 0 Then
                            .Range(1) = sOrderNumber & "-" & iCurrent
                        ElseIf InStr(sPantherModel, "MNS") <> 0 Then
                            sSerialNumber = Application.InputBox("Please Enter Serial Number of Printer")
                            .Range(1) = sSerialNumber
                        ElseIf InStr(sPantherModel, "P5C") <> 0 Then
                            sSerialNumber = Application.InputBox("Please Enter Serial Number on Wiring Plate")
                            .Range(1) = sSerialNumber
                        Else
                            .Range(1) = Application.WorksheetFunction.Max(loTable.ListColumns("Serial Number").Range) + 1 'serial number
                        End If
                        .Range(2) = Date             'date
                        .Range(3) = sOrderNumber     'order
                        .Range(4) = sCustomer        'customer
                        .Range(5) = sEndUser         'user
                        .Range(6) = sPantherModel    'model
                        .Range(7) = sLabelSize       'label size
                        
                        sSerialNumber = .Range(1)
                        
                        'Get ZPL Strings for Toe Tag Label
                        Call Get_Label_Strings(sPrinterName)
                        
                        sLabelZPL = "^XA^LH0,0^CI0^FD" & sLabelPart1 & sCustomer & sLabelPart2 & sEndUser & sLabelPart3 & sPantherModel + sLabelPart4 & sLabelSize _
                                    & sLabelPart5 & CStr(iCurrent) & " of " & CStr(iNumOrdered) & sLabelPart6 & sOrderNumber & sLabelPart7 & IIf(InStr(sPantherModel, "SLIDE") = 0, sSerialNumber & "^FS^PQ2^XZ", "^FS^PQ2^XZ")
                
                        'Check value of checkboxes and print out corresponding tags
                        Get_Tag_Checkbox
                        
                        If sModelName = "Shadow" Then
                            .Range(8) = sServoFirmware & sServoExt 'servo program
                        ElseIf sModelName = "STAND" Then
                            .Range(8) = sSystemFirmware & sSystemExt 'system program
                        ElseIf sModelName = "MNS" Then
                            .Range(8) = sServoFirmware 'servo program
                        Else
                            .Range(8) = sPLCFirmware & IIf(Len(sCustomFirmware) > 0 And shForm.CheckBoxes("PLC Checkbox").Value = 1, sCustomFirmware, "") & sPantherOptions & sPLCFileExt 'plc program
                            .Range(9) = sHMIFirmware & IIf(Len(sCustomFirmware) > 0 And shForm.CheckBoxes("HMI Checkbox").Value = 1, sCustomFirmware, "") & sHMIFileExt 'hmi program
                        End If
                    End With
                ElseIf InStr(sPantherModel, "SLIDE") <> 0 Then
                    Print_Tags
                End If
                
                
            Next iCurrent
        Next i
        shForm.Range("I5") = ""
        If firmwareExists = True Then
            MsgBox ("Serial Numbers Generated And Printed")
        End If
        
        'Save Entire Document
        Save_Document
        End
    End If
End Sub

Sub Duplicate_Search(sOrderNumber)
    Dim rgFound     As Range
    
    Set rgFound = ThisWorkbook.Sheets("P9").Range("C:C").Find(sOrderNumber)
    If rgFound Is Nothing Then
        Set rgFound = ThisWorkbook.Sheets("P5c").Range("C:C").Find(sOrderNumber)
        If rgFound Is Nothing Then
            Set rgFound = ThisWorkbook.Sheets("FLEX").Range("C:C").Find(sOrderNumber)
            If rgFound Is Nothing Then
                Exit Sub
            Else
                MsgBox ("Order Number already exists in FLEX : " & rgFound.Address)
            End If
        Else
            MsgBox ("Order Number already exists in P5c : " & rgFound.Address)
        End If
    Else
        MsgBox ("Order Number already exists in P9 : " & rgFound.Address)
    End If
    
End Sub

Sub Print_Machine_Tags(sZPL As String, sPrinterIP As String)
    Dim oHttp       As Object
    Dim sURL        As String
    Dim sZPLlength  As String
    
    Set oHttp = CreateObject("MSXML2.serverXMLHTTP")
    sURL = "http://" & sPrinterIP & "/pstprnt"
    sZPLlength = CStr(Len(sZPL))
    
    oHttp.Open "POST", sURL, True
    oHttp.setRequestHeader "Content-Length", sZPLlength
    oHttp.send sZPL
    Application.Wait Now + TimeValue("00:00:01")
End Sub

Sub Test_Print()
    Dim oHttp       As Object
    Dim sPrinterIP  As String
    Dim sURL        As String
    Dim sZPL        As String
    
    sPrinterIP = "192.168.17.97"
    
    Set oHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    sURL = "http://" & sPrinterIP & "/pstprnt"
    sZPL = "^XA^LH0,0^FT220,69^CI0^A0N,34,46^FDTest Customer^FS^FT220,121^A0N,34,46^FDTest End User _" & _
           "^FS^FO20,20^BY2,2,15^BQN,2,8^FDLA,12345^FS^FT45,240^A0N,34,46^FD12345^FS^FT48,284^A0N,34,46^FD1^FS _" & _
           "^FT70,284^A0N,34,46^FD of 2^FS^FT237,228^A0N,34,46^FDPA-P9-118e-4x6^FS^FT275,275^A0N,34,46^FDC000123456^FS^PQ1,0,0,Y^XZ"
    
    oHttp.Open "POST", sURL, True
    oHttp.setRequestHeader "Content-Length", CStr(Len(sZPL))
    oHttp.send sZPL
End Sub

Sub Print_Tags()
    
    'Variables
    Dim i          As Integer
    Dim sSlideHand As String
    
    'Target sheets
    Set shForm = ThisWorkbook.Sheets("Form")
    
    shForm.Range("I5") = "Printing"
    
    'Count number of rows with machines on order
    iMachinesOrdered = Application.WorksheetFunction.CountIf(shForm.Range("G12:G17"), "*")
    If iMachinesOrdered = 0 Then
        MsgBox ("There Are No Machines To Print Tags For")
        shForm.Range("I5") = ""
        End
    Else
        sOrderNumber = shForm.Range("G5")
        sCustomer = shForm.Range("G7")
        sEndUser = shForm.Range("G9")
        sPrinterName = shForm.Range("G19")
        
        For i = 1 To iMachinesOrdered

            bSlideOptions = False
            
            iMachineRow = 11 + i
            
            sPantherModel = shForm.Range("G" & iMachineRow).Value
            sPantherOptions = shForm.Range("J" & iMachineRow).Value
            
            'Check dictionary for model
            Firmware_Dictionary (sPantherModel)
            'Check To See If Options Work with the given model
            Validate_Options_Input
            If firmwareExists = False And InStr(sPantherModel, "SLIDE") = 0 Then
                addNewModelForm.Show
            End If
            If IsNumeric(shForm.Range("H" & iMachineRow)) = False Then
                MsgBox sPantherModel + " Qty Must Be a Number!"
                shForm.Range("I5") = ""
                End
            ElseIf shForm.Range("H" & iMachineRow) < 1 Then
                MsgBox sPantherModel + " Qty Must Be Greater Than 0!"
                shForm.Range("I5") = ""
                End
            Else
                iNumOrdered = shForm.Range("H" & iMachineRow)
            End If
            
            sLabelSize = shForm.Range("I" & iMachineRow)
                     
            If InStr(sPantherModel, "SLIDE") = 0 And InStr(sPantherModel, "STAND") = 0 And InStr(sPantherModel, "MNS") = 0 Then
                sSerialNumber = Application.InputBox("Enter Starting Serial Number For " & sPantherModel, "Serial Numbers")
                If sSerialNumber = "" Then
                    shForm.Range("I5") = ""
                    MsgBox "Serial Number Can't Be Blank!"
                    End
                ElseIf sSerialNumber = False Then
                    shForm.Range("I5") = ""
                    End
                End If
            End If
            
            If InStr(sPantherModel, "STAND") <> 0 Then
                sSerialNumber = sOrderNumber
            End If
            
            For iCurrent = 1 To iNumOrdered
            
                If InStr(sPantherModel, "MNS") <> 0 Then
                    sSerialNumber = Application.InputBox("Please Enter Serial Number of Printer")
                End If
                
                'Get ZPL Strings for Toe Tag Label
                Call Get_Label_Strings(sPrinterName)
            
                If InStr(sPantherModel, "SLIDE") <> 0 Then
                    sSlideHand = shForm.Range("J" & iMachineRow)
                    If bSlideOptions = False Then
                        sSlideHand = Application.InputBox("Please enter hand for " & IntToOrdinalString(iCurrent) & " " & sPantherModel & " slide")
                    End If
                    sLabelZPL = "^XA^LH0,0^CI0 ^FD" & sLabelPart1 & sCustomer & sLabelPart2 & sEndUser & sLabelPart3 & sPantherModel + sLabelPart4 & sSlideHand _
                                & sLabelPart5 & CStr(iCurrent) & " of " & CStr(iNumOrdered) & sLabelPart6 & sOrderNumber & sLabelPart7 _
                                & "^FS^PQ1^XZ"
                ElseIf InStr(sPantherModel, "STAND") <> 0 Then
                    sLabelZPL = "^XA^LH0,0^CI0 ^FD" & sLabelPart1 & sCustomer & sLabelPart2 & sEndUser & sLabelPart3 & sPantherModel + sLabelPart4 & sLabelSize _
                                & sLabelPart5 & CStr(iCurrent) & " of " & CStr(iNumOrdered) & sLabelPart6 & sOrderNumber & sLabelPart7 _
                                & sSerialNumber & "-" & iCurrent & "^FS^PQ2^XZ"
                Else
                    sLabelZPL = "^XA^LH0,0^CI0 ^FD" & sLabelPart1 & sCustomer & sLabelPart2 & sEndUser & sLabelPart3 & sPantherModel + sLabelPart4 & sLabelSize _
                                & sLabelPart5 & CStr(iCurrent) & " of " & CStr(iNumOrdered) & sLabelPart6 & sOrderNumber & sLabelPart7 _
                                & sSerialNumber & "^FS^PQ2^XZ"
                End If
                
                'Check value of checkboxes and print out corresponding tags
                Get_Tag_Checkbox
                
                If InStr(sPantherModel, "SLIDE") = 0 And InStr(sPantherModel, "STAND") = 0 And InStr(sPantherModel, "MNS") = 0 Then
                    sSerialNumber = sSerialNumber + 1
                End If
                
            Next iCurrent
            
        Next i
        shForm.Range("I5") = ""
        If firmwareExists = True Then
            MsgBox ("Serial Numbers Printed")
            Save_Document
        End If
        
    End If
    
End Sub

Sub Save_Document()

Dim wb      As Workbook

    'Loop through each open workbook and save it
    'Basically autosave since excel can only autosave to onedrive instead of locally
    For Each wb In Workbooks
        wb.Save
    Next wb

End Sub

Sub Validate_Options_Input()
    'Check Options field and Compare with the Model Name to see if combination is possible
    If (sPantherOptions = "AE") Then
        If (sModelName = "Flex") Or (sModelName = "P5c") Or (sModelName = "Shadow") Or (sModelName = "STAND") Or (sModelName = "Predator Straight Tamp (Beijer)") _
        Or InStr(sPantherModel, "SLIDE") <> 0 Then
            MsgBox "A " & sModelName & " Machine Can't Have Auto Height and or Expansion as an Option."
            shForm.Range("I5") = ""
            End
        End If
    ElseIf sPantherOptions = "A" Then
        If (sModelName = "Predator Swing Arm") Or (sModelName = "Phantom") Or (sModelName = "P5c") _
        Or (sModelName = "Flex") Or (sModelName = "Shadow") Or (sModelName = "STAND") Or (sModelName = "Predator Straight Tamp (Beijer)") _
        Or InStr(sPantherModel, "SLIDE") <> 0 Then
            MsgBox "A " & sModelName & " Machine Can't Have Auto Height as an Option."
            shForm.Range("I5") = ""
            End
        End If
    ElseIf sPantherOptions = "E" Then
        If (sModelName = "Flex") Or (sModelName = "P5c") _
        Or (sModelName = "Shadow") Or (sModelName = "STAND") Or InStr(sPantherModel, "SLIDE") <> 0 Then
            MsgBox "A " & sModelName & " Machine Can't Have Expansion as an Option."
            shForm.Range("I5") = ""
            End
        End If
    ElseIf sPantherOptions <> "" Then
        If InStr(sPantherModel, "SLIDE") <> 0 Then
            If (sPantherOptions = "RH") Or (sPantherOptions = "LH") Then
                bSlideOptions = True
            Else
                MsgBox "SLIDE Hand Must be Either RH or LH"
                shForm.Range("I5") = ""
                End
                End If
        ElseIf sPantherOptions <> "A" Then
            MsgBox "Options Must Have Either A for AutoHeight, E for Expansion, or AE for Both."
            shForm.Range("I5") = ""
            End
        ElseIf sPantherOptions <> "E" Then
            MsgBox "Options Must Have Either A for AutoHeight, E for Expansion, or AE for Both."
            shForm.Range("I5") = ""
            End
        ElseIf sPantherOptions <> "AE" Then
            MsgBox "Options Must Have Either A for AutoHeight, E for Expansion, or AE for Both."
            shForm.Range("I5") = ""
            End
        End If
    End If
    
End Sub

Public Function IntToOrdinalString(MyNumber As Integer) As String
    Dim sOutput As String
    Dim iUnit As Integer
    
    
    iUnit = MyNumber Mod 10
    sOutput = ""
    
    Select Case MyNumber
        Case Is < 0
            sOutput = ""
        Case 10 To 19
            sOutput = "th"
        Case Else
            Select Case iUnit
                Case 0 'Zeroth only has a meaning when counts start with zero, which happens in a mathematical or computer science context.
                    sOutput = "th"
                Case 1
                    sOutput = "st"
                Case 2
                    sOutput = "nd"
                Case 3
                    sOutput = "rd"
                Case 4 To 9
                    sOutput = "th"
            End Select
    End Select
    IntToOrdinalString = CStr(MyNumber) & sOutput
End Function
