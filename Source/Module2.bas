Attribute VB_Name = "Connections"
Public Sub ShowSupplierData(SQLString As String)
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.CursorLocation = adUseClient
    rsSupplier.CursorType = adOpenStatic
    rsSupplier.LockType = adLockReadOnly
    rsSupplier.Open SQLString, conn
        If rsSupplier.EOF = True Then
            rsSupplier.Close
            Set rsSupplier = Nothing
            Exit Sub
        End If
    xCount = rsSupplier.RecordCount
        If Rx > rsSupplier.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsSupplier.RecordCount - 1
        End If
    rsSupplier.Move Rx
    frmSupplier.txtSID.Text = rsSupplier!Supplier_ID
    frmSupplier.txtDate.Text = Format(rsSupplier!Date, "YYYY-MM-DD")
    frmSupplier.txtSName.Text = rsSupplier!Name
    frmSupplier.txtCompany.Text = rsSupplier!Company
    frmSupplier.txtCP.Text = rsSupplier!Contact_Person
    frmSupplier.txtAddress.Text = rsSupplier!Address
    frmSupplier.txtOfficeNo.Text = rsSupplier!Office_No
    frmSupplier.txtMobileNo.Text = rsSupplier!Mobile_No
    frmSupplier.txtOtherNo.Text = rsSupplier!Other_No
    frmSupplier.txtFaxNo.Text = rsSupplier!Fax_No
    frmSupplier.txtTBM.Text = rsSupplier!Total_Bills_Amount
    frmSupplier.txtDue.Text = rsSupplier!Total_Due
    frmSupplier.txtR.Text = rsSupplier!Remarks
    
    rsSupplier.Close
    Set rsSupplier = Nothing
End Sub
Public Sub ShowSupplierGrid(SQLString As String)
    Set RsSuppGrid = New ADODB.Recordset
    RsSuppGrid.CursorLocation = adUseClient
    RsSuppGrid.CursorType = adOpenStatic
    RsSuppGrid.LockType = adLockReadOnly
    RsSuppGrid.Open SQLString, conn
    Set frmSupplier.DataGrid1.DataSource = RsSuppGrid
End Sub

Public Sub ShowSupplierAccountData(SQLString As String)
    Set rsSupplierAccount = New ADODB.Recordset
    rsSupplierAccount.CursorLocation = adUseClient
    rsSupplierAccount.CursorType = adOpenStatic
    rsSupplierAccount.LockType = adLockReadOnly
    rsSupplierAccount.Open SQLString, conn
        If rsSupplierAccount.EOF = True Then
            rsSupplierAccount.Close
            Set rsSupplierAccount = Nothing
            Exit Sub
        End If
    xCount = rsSupplierAccount.RecordCount
        If Rx > rsSupplierAccount.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsSupplierAccount.RecordCount - 1
        End If
    rsSupplierAccount.Move Rx
    frmSupplierAccount.txtTID = rsSupplierAccount!TID
    frmSupplierAccount.txtDate.Text = Format(rsSupplierAccount!Date, "YYYY-MM-DD")
    frmSupplierAccount.txtSID.Text = rsSupplierAccount!Supplier_ID
    frmSupplierAccount.txtPO.Text = rsSupplierAccount!PO_No
    frmSupplierAccount.txtTA.Text = rsSupplierAccount!Total_Amount
    frmSupplierAccount.PM.Text = rsSupplierAccount!Payment_Mode
    frmSupplierAccount.txtPA.Text = rsSupplierAccount!Paid_Amount
    frmSupplierAccount.txtDA.Text = rsSupplierAccount!Due_Amount
    frmSupplierAccount.txtR.Text = rsSupplierAccount!Remarks
    
    rsSupplierAccount.Close
    Set rsSupplierAccount = Nothing
End Sub
Public Sub ShowSupplierAccountGrid(SQLString As String)
    Set RsSuppAccountGrid = New ADODB.Recordset
    RsSuppAccountGrid.CursorLocation = adUseClient
    RsSuppAccountGrid.CursorType = adOpenStatic
    RsSuppAccountGrid.LockType = adLockReadOnly
    RsSuppAccountGrid.Open SQLString, conn
    Set frmSupplierAccount.DataGrid1.DataSource = RsSuppAccountGrid
End Sub

Public Sub ShowPOData(SQLString As String)
    Set rsPO = New ADODB.Recordset
    rsPO.CursorLocation = adUseClient
    rsPO.CursorType = adOpenStatic
    rsPO.LockType = adLockReadOnly
    rsPO.Open SQLString, conn
        If rsPO.EOF = True Then
            rsPO.Close
            Set rsPO = Nothing
            Exit Sub
        End If
    xCount = rsPO.RecordCount
        If Rx > rsPO.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsPO.RecordCount - 1
        End If
    rsPO.Move Rx
    frmPurchaseOrder.txtPO = rsPO!PO_No
    frmPurchaseOrder.txtDate.Text = Format(rsPO!Date, "YYYY-MM-DD")
    frmPurchaseOrder.txtSID.Text = rsPO!Supplier_ID
    frmPurchaseOrder.txtDD.Text = Format(rsPO!Date, "YYYY-MM-DD")
    frmPurchaseOrder.txtR.Text = rsPO!Remarks
'    frmPurchaseOrder.Prod.Text = rsPO!Product
'    frmPurchaseOrder.txtPT.Text = rsPO!Product_Type
'    frmPurchaseOrder.txtQty.Text = rsPO!Quantity
'    frmPurchaseOrder.txtSize.Text = rsPO!Size
'    frmPurchaseOrder.txtDescription.Text = rsPO!Description
    
    rsPO.Close
    Set rsPO = Nothing
End Sub
Public Sub ShowPOGrid(SQLString As String)
    Set RsPOGrid = New ADODB.Recordset
    RsPOGrid.CursorLocation = adUseClient
    RsPOGrid.CursorType = adOpenStatic
    RsPOGrid.LockType = adLockReadOnly
    RsPOGrid.Open SQLString, conn
    Set frmPurchaseOrder.DataGrid1.DataSource = RsPOGrid
End Sub

Public Sub ShowStockData(SQLString As String)
    Set rsStock = New ADODB.Recordset
    rsStock.CursorLocation = adUseClient
    rsStock.CursorType = adOpenStatic
    rsStock.LockType = adLockReadOnly
    rsStock.Open SQLString, conn
        If rsStock.EOF = True Then
            rsStock.Close
            Set rsStock = Nothing
            Exit Sub
        End If
    xCount = rsStock.RecordCount
        If Rx > rsStock.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsStock.RecordCount - 1
        End If
    rsStock.Move Rx
    frmStock.txtPID = rsStock!Product_ID
    frmStock.txtDate.Text = Format(rsStock!Date, "YYYY-MM-DD")
    frmStock.txtProduct.Text = rsStock!Product
    frmStock.PType.Text = rsStock!Product_Type
    frmStock.txtPS.Text = rsStock!Product_Size
    frmStock.Company.Text = rsStock!Company
    frmStock.txtBuyPrice.Text = rsStock!buying_Price
    frmStock.txtSelPrice.Text = rsStock!selling_Price
    frmStock.txtDescription.Text = rsStock!Description
    frmStock.txtStock.Text = rsStock!Stock_In_Hand
    frmStock.txtROL.Text = rsStock!ReOrder_Level
    frmStock.txtR.Text = rsStock!Remarks
    
    rsStock.Close
    Set rsStock = Nothing
End Sub
Public Sub ShowStockGrid(SQLString As String)
    Set rsStockGrid = New ADODB.Recordset
    rsStockGrid.CursorLocation = adUseClient
    rsStockGrid.CursorType = adOpenStatic
    rsStockGrid.LockType = adLockReadOnly
    rsStockGrid.Open SQLString, conn
    Set frmStock.DataGrid1.DataSource = rsStockGrid
End Sub

Public Sub ShowReceivingsData(SQLString As String)
    Set rsReceivings = New ADODB.Recordset
    rsReceivings.CursorLocation = adUseClient
    rsReceivings.CursorType = adOpenStatic
    rsReceivings.LockType = adLockReadOnly
    rsReceivings.Open SQLString, conn
        If rsReceivings.EOF = True Then
            rsReceivings.Close
            Set rsReceivings = Nothing
            Exit Sub
        End If
    xCount = rsReceivings.RecordCount
        If Rx > rsReceivings.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsReceivings.RecordCount - 1
        End If
    rsReceivings.Move Rx
    frmReceivings.txtTID = rsReceivings!TID
    frmReceivings.txtDate.Text = Format(rsReceivings!Date, "YYYY-MM-DD")
    frmReceivings.txtPO = rsReceivings!PO_No
    frmReceivings.txtPID.Text = rsReceivings!Product_ID
    frmReceivings.txtQty.Text = rsReceivings!Quantity
    frmReceivings.txtPrice.Text = rsReceivings!Price
    frmReceivings.txtPricePU.Text = rsReceivings!Price_Per_Unit
    frmReceivings.txtR.Text = rsReceivings!Remarks
    
    rsReceivings.Close
    Set rsReceivings = Nothing
End Sub
Public Sub ShowReceivingsGrid(SQLString As String)
    Set RsReceivingsGrid = New ADODB.Recordset
    RsReceivingsGrid.CursorLocation = adUseClient
    RsReceivingsGrid.CursorType = adOpenStatic
    RsReceivingsGrid.LockType = adLockReadOnly
    RsReceivingsGrid.Open SQLString, conn
    Set frmReceivings.DataGrid1.DataSource = RsReceivingsGrid
End Sub

Public Sub ShowCustomerData(SQLString As String)
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.CursorLocation = adUseClient
    rsCustomer.CursorType = adOpenStatic
    rsCustomer.LockType = adLockReadOnly
    rsCustomer.Open SQLString, conn
        If rsCustomer.EOF = True Then
            rsCustomer.Close
            Set rsCustomer = Nothing
            Exit Sub
        End If
    xCount = rsCustomer.RecordCount
        If Rx > rsCustomer.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsCustomer.RecordCount - 1
        End If
    rsCustomer.Move Rx
    frmCustomer.txtCID.Text = rsCustomer!Customer_ID
    frmCustomer.txtDate.Text = Format(rsCustomer!Date, "YYYY-MM-DD")
    frmCustomer.txtName.Text = rsCustomer!Name
    frmCustomer.txtCNIC.Text = rsCustomer!CNIC_No
    frmCustomer.txtAddress.Text = rsCustomer!Address
    frmCustomer.txtOCP.Text = rsCustomer!Occupation
    frmCustomer.txtPhone.Text = rsCustomer!Phone_No
    frmCustomer.txtMobileNo.Text = rsCustomer!Mobile_No
    frmCustomer.txtOtherNo.Text = rsCustomer!Other_No
    frmCustomer.txtTBM.Text = rsCustomer!Total_Bills_Amount
    frmCustomer.txtDue.Text = rsCustomer!Total_Due
    frmCustomer.txtR.Text = rsCustomer!Remarks
    
    rsCustomer.Close
    Set rsCustomer = Nothing
End Sub
Public Sub ShowCustomerGrid(SQLString As String)
    Set RsCustomerGrid = New ADODB.Recordset
    RsCustomerGrid.CursorLocation = adUseClient
    RsCustomerGrid.CursorType = adOpenStatic
    RsCustomerGrid.LockType = adLockReadOnly
    RsCustomerGrid.Open SQLString, conn
    Set frmCustomer.DataGrid1.DataSource = RsCustomerGrid
End Sub

Public Sub ShowCustomerAccountData(SQLString As String)
    Set rsCustomerAccount = New ADODB.Recordset
    rsCustomerAccount.CursorLocation = adUseClient
    rsCustomerAccount.CursorType = adOpenStatic
    rsCustomerAccount.LockType = adLockReadOnly
    rsCustomerAccount.Open SQLString, conn
        If rsCustomerAccount.EOF = True Then
            rsCustomerAccount.Close
            Set rsCustomerAccount = Nothing
            Exit Sub
        End If
    xCount = rsCustomerAccount.RecordCount
        If Rx > rsCustomerAccount.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsCustomerAccount.RecordCount - 1
        End If
    rsCustomerAccount.Move Rx
    frmCustomerAccount.txtTID = rsCustomerAccount!TID
    frmCustomerAccount.txtDate.Text = Format(rsCustomerAccount!Date, "YYYY-MM-DD")
    frmCustomerAccount.txtCID.Text = rsCustomerAccount!Customer_ID
    frmCustomerAccount.txtInv.Text = rsCustomerAccount!Invoice_No
    frmCustomerAccount.txtTA.Text = rsCustomerAccount!Total_Amount
    frmCustomerAccount.PM.Text = rsCustomerAccount!Payment_Mode
    frmCustomerAccount.txtPA.Text = rsCustomerAccount!Amount_Paid
    frmCustomerAccount.txtDA.Text = rsCustomerAccount!Amount_Due
    frmCustomerAccount.txtR.Text = rsCustomerAccount!Remarks
    
    rsCustomerAccount.Close
    Set rsCustomerAccount = Nothing
End Sub
Public Sub ShowCustomerAccountGrid(SQLString As String)
    Set RsCustomerAccountGrid = New ADODB.Recordset
    RsCustomerAccountGrid.CursorLocation = adUseClient
    RsCustomerAccountGrid.CursorType = adOpenStatic
    RsCustomerAccountGrid.LockType = adLockReadOnly
    RsCustomerAccountGrid.Open SQLString, conn
    Set frmCustomerAccount.DataGrid1.DataSource = RsCustomerAccountGrid
End Sub

Public Sub ShowInvoiceData(SQLString As String)
    Set rsInvoice = New ADODB.Recordset
    rsInvoice.CursorLocation = adUseClient
    rsInvoice.CursorType = adOpenStatic
    rsInvoice.LockType = adLockReadOnly
    rsInvoice.Open SQLString, conn
        If rsInvoice.EOF = True Then
            rsInvoice.Close
            Set rsInvoice = Nothing
            Exit Sub
        End If
    xCount = rsInvoice.RecordCount
        If Rx > rsInvoice.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsInvoice.RecordCount - 1
        End If
    rsInvoice.Move Rx
    frmInvoice.txtInv = rsInvoice!Invoice_No
    frmInvoice.txtDate.Text = Format(rsInvoice!Date, "YYYY-MM-DD")
    frmInvoice.txtSalesman.Text = rsInvoice!Salesman
    frmInvoice.txtCID = rsInvoice!Customer_ID
    frmInvoice.txtGT.Text = rsInvoice!Grand_Total
    frmInvoice.Discount.Text = rsInvoice!Discount
    frmInvoice.PM.Text = rsInvoice!Payment_Mode
    frmInvoice.txtAP.Text = rsInvoice!Amount_Paid
    frmInvoice.txtChange.Text = rsInvoice!Amount_Change
    frmInvoice.txtAD.Text = rsInvoice!Amount_Due
    frmInvoice.txtProfit.Text = rsInvoice!Profit
    frmInvoice.txtR.Text = rsInvoice!Remarks
    
    rsInvoice.Close
    Set rsInvoice = Nothing
End Sub
Public Sub ShowInvoiceGrid(SQLString As String)
    Set RsInvoiceGrid = New ADODB.Recordset
    RsInvoiceGrid.CursorLocation = adUseClient
    RsInvoiceGrid.CursorType = adOpenStatic
    RsInvoiceGrid.LockType = adLockReadOnly
    RsInvoiceGrid.Open SQLString, conn
    Set frmInvoice.DataGrid1.DataSource = RsInvoiceGrid
End Sub

Public Sub ShowUserData(SQLString As String)
    Set RsUser = New ADODB.Recordset
    RsUser.CursorLocation = adUseClient
    RsUser.CursorType = adOpenStatic
    RsUser.LockType = adLockReadOnly
    RsUser.Open SQLString, conn
        If RsUser.EOF = True Then
            RsUser.Close
            Set RsUser = Nothing
            Exit Sub
        End If
    xCount = RsUser.RecordCount
        If Rx > RsUser.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = RsUser.RecordCount - 1
        End If
    
    RsUser.Move Rx
    frmSecurity.txtUser.Text = RsUser!user
    frmSecurity.txtPass.Text = RsUser!Password
    frmSecurity.UType.Text = RsUser!Type
    frmSecurity.txtFN.Text = RsUser!Name
    frmSecurity.txtDesg.Text = RsUser!Designation
    frmSecurity.txtR.Text = RsUser!Remarks
    
    RsUser.Close
    Set RsUser = Nothing
End Sub
Public Sub ShowUserGrid()
    Set RsUserGrid = New ADODB.Recordset
    RsUserGrid.CursorLocation = adUseClient
    RsUserGrid.CursorType = adOpenStatic
    RsUserGrid.LockType = adLockReadOnly
    RsUserGrid.Open SQLString, conn
    Set frmSecurity.DataGrid1.DataSource = RsUserGrid
End Sub

Public Sub ShowExpenseData(SQLString As String)
    Set rsExpense = New ADODB.Recordset
    rsExpense.CursorLocation = adUseClient
    rsExpense.CursorType = adOpenStatic
    rsExpense.LockType = adLockReadOnly
    rsExpense.Open SQLString, conn
        If rsExpense.EOF = True Then
            rsExpense.Close
            Set rsExpense = Nothing
            Exit Sub
        End If
    xCount = rsExpense.RecordCount
        If Rx > rsExpense.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsExpense.RecordCount - 1
        End If
    rsExpense.Move Rx
    frmExpense.txtTID.Text = rsExpense!TID
    frmExpense.txtDate.Text = Format(rsExpense!Date, "YYYY-MM-DD")
    frmExpense.ET.Text = rsExpense!Expense_Type
    frmExpense.txtSupplier.Text = rsExpense!Supplier
    frmExpense.PM.Text = rsExpense!Payment_Mode
    frmExpense.txtParticulars.Text = rsExpense!Particulars
    frmExpense.txtAmount.Text = rsExpense!Amount
    frmExpense.txtR.Text = rsExpense!Remarks
    
    rsExpense.Close
    Set rsExpense = Nothing
End Sub
Public Sub ShowExpenseGrid(SQLString As String)
    Set RsExpenseGrid = New ADODB.Recordset
    RsExpenseGrid.CursorLocation = adUseClient
    RsExpenseGrid.CursorType = adOpenStatic
    RsExpenseGrid.LockType = adLockReadOnly
    RsExpenseGrid.Open SQLString, conn
    Set frmExpense.DataGrid1.DataSource = RsExpenseGrid
End Sub
