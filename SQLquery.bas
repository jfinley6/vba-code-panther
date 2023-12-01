Attribute VB_Name = "SQLquery"
Option Explicit

Sub SQL_Query()
    
    Dim shForm As Worksheet
    Set shForm = ThisWorkbook.Sheets("Form")
    
    Dim sMacroInUse As String
    sMacroInUse = shForm.Range("I5")
        
    If sMacroInUse = "Searching" Then
        MsgBox ("Search is Busy, please wait and try again")
        Exit Sub
    End If
        
    shForm.Range("I5") = "Searching"
    
    'ThisWorkbook.Sheets("Form").Range("G7, G9, G12:I17").Value = ""

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = "DSN=SL_Reporting;Description=SL_Reporting;UID=SL_Reporting;Trusted_Connection=Yes;APP=Microsoft Office;DATABASE=IDT_App;"
    cn.Open

    Dim sSQLQueryCustomer As String
    Dim sSQLQueryEndUser As String
    Dim sSQLQueryLineItems As String
    
    Dim sOrderNumber As String
    Dim sCustomer As String
    Dim sCustomerName As String
    Dim sEndUser As String
    Dim sPantherModel As String
    
    
    sOrderNumber = shForm.Range("G5")
    
    sSQLQueryCustomer = "SELECT Uf_DUNSNAME FROM customer_mst INNER JOIN co_mst ON customer_mst.cust_num = co_mst.cust_num AND customer_mst.cust_seq = 0 WHERE co_num='" & sOrderNumber & "'"
    Set rs = cn.Execute(sSQLQueryCustomer)
    
    If rs.EOF Then
        shForm.Range("J5") = ""
        MsgBox ("There are no matching orders")
        Exit Sub
    End If
    
    sCustomer = rs.Fields("Uf_DUNSNAME").Value
 
    sSQLQueryEndUser = "SELECT Uf_DUNSNAME FROM customer_mst INNER JOIN co_mst ON customer_mst.cust_num = co_mst.cust_num AND customer_mst.cust_seq = co_mst.cust_seq WHERE co_num='" & sOrderNumber & "'"
    Set rs = cn.Execute(sSQLQueryEndUser)
    sEndUser = rs.Fields("Uf_DUNSNAME").Value
    
    rs.Close
    Set rs = Nothing
        
    sSQLQueryLineItems = "SELECT item, qty_ordered FROM coitem_mst WHERE whse = 'HRA' AND item LIKE 'PA-%' AND price > 5000 AND co_num='" & sOrderNumber & "';"
    Set rs = cn.Execute(sSQLQueryLineItems)
    
    If rs.EOF Then
        shForm.Range("J5") = ""
        MsgBox ("There are no Panther machines on this order")
        Exit Sub
    End If
    
    shForm.Range("G12:G17").CopyFromRecordset rs
    
    rs.Close
    Set rs = Nothing
    
    shForm.Range("G7") = sCustomer
    shForm.Range("G9") = sEndUser
    shForm.Range("G11") = sPantherModel
    
    cn.Close
    
    shForm.Range("I5") = ""
    
End Sub

