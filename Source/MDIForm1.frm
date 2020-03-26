VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   Caption         =   ":: All About Innovations :.                             .: MEGA SHOPPING MALL :.                             :: Point of Sale :."
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0ECA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   8625
      Width           =   13680
      _ExtentX        =   24130
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   13
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "MDIForm1.frx":1A9B1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "User Name:"
            TextSave        =   "User Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "MDIForm1.frx":1AF4B
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Time Log-in:"
            TextSave        =   "Time Log-in:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "MDIForm1.frx":1B2E5
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/2/2007"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnClose 
         Caption         =   "&Close"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnFS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnCB 
         Caption         =   "Close Batch"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnPurchase 
      Caption         =   "&Purchase"
      Begin VB.Menu mnSupplier 
         Caption         =   "Suppliers Data"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnSuppAcc 
         Caption         =   "Supplier Accounts"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnPO 
         Caption         =   "Purchase Order"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnSM 
      Caption         =   "&Stock Management"
      Begin VB.Menu mnReceivings 
         Caption         =   "Receivings"
         Shortcut        =   ^R
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnInventory 
         Caption         =   "Stock Inventory"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnSales 
      Caption         =   "&Sales"
      Begin VB.Menu mnCustomer 
         Caption         =   "Customer Data"
         Shortcut        =   ^C
         Visible         =   0   'False
      End
      Begin VB.Menu mnCusAcc 
         Caption         =   "Customer Accounts"
         Visible         =   0   'False
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnInvoice 
         Caption         =   "Invoice"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnGeneral 
      Caption         =   "&General"
      Begin VB.Menu mnGE 
         Caption         =   "General Expenditure"
      End
   End
   Begin VB.Menu mnReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnRptS 
         Caption         =   "Database Summary"
         Shortcut        =   ^D
      End
      Begin VB.Menu Sep007 
         Caption         =   "-"
      End
      Begin VB.Menu mnSupplierR 
         Caption         =   "Supplier"
         Begin VB.Menu mnRptSupp 
            Caption         =   "Suppliers Data"
         End
         Begin VB.Menu mnRptSuppAcc 
            Caption         =   "Supplier Accounts"
         End
         Begin VB.Menu mnRptDueSupp 
            Caption         =   "Due Supplier Accounts"
         End
         Begin VB.Menu mnRptOKSupp 
            Caption         =   "Cleared Supplier Accounts"
         End
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnCustomerR 
         Caption         =   "Customer"
         Visible         =   0   'False
         Begin VB.Menu mnRptCus 
            Caption         =   "Customers Data"
         End
         Begin VB.Menu mnRptCusAcc 
            Caption         =   "Customer Accounts"
         End
         Begin VB.Menu mnRptDueCus 
            Caption         =   "Due Customer Accounts"
         End
         Begin VB.Menu mnRptOKCus 
            Caption         =   "Cleared Customer Accounts"
         End
      End
      Begin VB.Menu Sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnRptPr 
         Caption         =   "Stock"
         Begin VB.Menu mnRptPO 
            Caption         =   "Particular Purchase Order"
         End
         Begin VB.Menu mnRptPOs 
            Caption         =   "Purchase Orders"
         End
         Begin VB.Menu mnRptSPOBD 
            Caption         =   "Purchase Orders Between Two Dates"
         End
         Begin VB.Menu SepSt1 
            Caption         =   "-"
         End
         Begin VB.Menu mnRptReceiving 
            Caption         =   "Receivings"
         End
         Begin VB.Menu mnRptSRBD 
            Caption         =   "Receivings Between Two Dates"
         End
         Begin VB.Menu mnRptTRec 
            Caption         =   "Today's Receivings"
            Visible         =   0   'False
         End
         Begin VB.Menu Sep5 
            Caption         =   "-"
         End
         Begin VB.Menu mnRptStock 
            Caption         =   "Total Stock"
         End
         Begin VB.Menu mnRptPS 
            Caption         =   "Particular Product Stock"
         End
      End
      Begin VB.Menu Sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnSaleR 
         Caption         =   "Sale"
         Begin VB.Menu mnRptInv 
            Caption         =   "Particular Invoice"
         End
         Begin VB.Menu mnRptSales 
            Caption         =   "Sales Report"
         End
         Begin VB.Menu mnRptTS 
            Caption         =   "Today's Sale"
         End
         Begin VB.Menu mnRptSBD 
            Caption         =   "Sale Between Two Dates"
         End
      End
      Begin VB.Menu Sep07 
         Caption         =   "-"
      End
      Begin VB.Menu mnProfit 
         Caption         =   "Profit"
         Begin VB.Menu mnRptProfit 
            Caption         =   "Profit Report"
         End
         Begin VB.Menu mnRptProfitChart 
            Caption         =   "Monthly Profit Chart [1 Year]"
         End
         Begin VB.Menu mnRptProfitYearChart 
            Caption         =   "Yearly Profit Chart"
         End
      End
      Begin VB.Menu Sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnRptG 
         Caption         =   "General"
         Begin VB.Menu mnRptGE 
            Caption         =   "General Expenditure"
         End
         Begin VB.Menu mnRptGEBD 
            Caption         =   "General Expenditure Between Two Dates"
         End
      End
   End
   Begin VB.Menu mnOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnChangePwd 
         Caption         =   "&Change Password"
         Shortcut        =   {F7}
      End
      Begin VB.Menu SepOpt 
         Caption         =   "-"
      End
      Begin VB.Menu mnUM 
         Caption         =   "&User Management"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnArrange 
         Caption         =   "Arrange"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnCascade 
         Caption         =   "Cascade"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnHorizontal 
         Caption         =   "Horizontal"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnVertical 
         Caption         =   "Vertical"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnAbout 
      Caption         =   "&About"
      Begin VB.Menu mnSoftware 
         Caption         =   "Software"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EndMe As Boolean
Dim RptStr, RptStr2 As String

Private Sub MDIForm_Load()
    CheckUser
    EndMe = True
    StatusBar1.Panels(3) = UserName
    StatusBar1.Panels(6) = LoginTime
    
    isStockMinus = False
    isReOrder = False
    


    'Me.Height = 9100
    'Me.Width = 12000
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Are you Sure you want to exit? ", vbYesNo + vbQuestion, "Quit?") = vbYes Then
        UpdateProfit
        BackupDatabase
        conn.Close
        Set conn = Nothing
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub mnCB_Click()
    UpdateProfit
    BackupDatabase
    MsgBox "Batch Closed & Database Backup Sucessfull", vbInformation, "POS"
End Sub

Private Sub mnClose_Click()
    If ActiveForm Is Nothing Then Exit Sub
    Unload ActiveForm
End Sub

Private Sub mnCusAcc_Click()
    frmCustomerAccount.Show
    frmCustomerAccount.Left = 20
    frmCustomerAccount.Top = 20
End Sub

Private Sub mnCustomer_Click()
    frmCustomer.Show
    frmCustomer.Left = 20
    frmCustomer.Top = 20
End Sub

Private Sub mnExit_Click()
    If MsgBox("Are you Sure you want to exit? ", vbYesNo + vbQuestion, "Quit?") = vbYes Then
        UpdateProfit
        BackupDatabase
        conn.Close
        Set conn = Nothing
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub mnGE_Click()
    frmExpense.Show
    frmExpense.Left = 20
    frmExpense.Top = 20
End Sub



Private Sub mnHorizontal_Click()
    MDIForm1.Arrange vbTileHorizontal
End Sub

Private Sub mnInventory_Click()
    frmStock.Show
    frmStock.Left = 20
    frmStock.Top = 20
End Sub

Private Sub mnInvoice_Click()
    'MsgBox "Invoice From will be operational very soon!", vbInformation
    frmInvoice.Show
    frmInvoice.Left = 20
    frmInvoice.Top = 20
End Sub

Private Sub mnPO_Click()
    frmPurchaseOrder.Show
    frmPurchaseOrder.Left = 20
    frmPurchaseOrder.Top = 20
End Sub

Private Sub mnReceivings_Click()
    frmReceivings.Show
    frmReceivings.Left = 20
    frmReceivings.Top = 20
End Sub

Private Sub mnRptCus_Click()
    RptCustomer.Show
End Sub

Private Sub mnRptCusAcc_Click()
    Unload RptCusAcct
    RptSql = "SELECT Customer_Account.TID, Customer_Account.Date, Customer.Name, Customer_Account.Invoice_No, Customer_Account.Total_Amount, Customer_Account.Payment_Mode, Customer_Account.Amount_Paid, Customer_Account.Amount_Due FROM Customer_Account,Customer WHERE Customer_Account.Customer_ID=Customer.Customer_ID ORDER BY Customer_Account.TID;"
    RptCusAcct.Show
End Sub

Private Sub mnRptDueCus_Click()
    Unload RptCusAcct
    RptSql = "SELECT Customer_Account.TID, Customer_Account.Date, Customer.Name, Customer_Account.Invoice_No, Customer_Account.Total_Amount, Customer_Account.Payment_Mode, Customer_Account.Amount_Paid, Customer_Account.Amount_Due FROM Customer_Account,Customer WHERE Customer_Account.Customer_ID=Customer.Customer_ID AND Customer_Account.Amount_Due<>0 ORDER BY Customer_Account.TID;"
    RptCusAcct.Show
End Sub

Private Sub mnRptDueSupp_Click()
    Unload RptSuppAcct
    RptSql = "SELECT Supplier_Account.TID, Supplier_Account.Date,Supplier.Company,Supplier_Account.PO_No,Supplier_Account.Total_Amount,Supplier_Account.Payment_Mode,Supplier_Account.Paid_Amount,Supplier_Account.Due_Amount FROM Supplier_Account,Supplier WHERE Supplier_Account.Supplier_ID=Supplier.Supplier_ID AND Supplier_Account.Due_Amount<>0 ORDER BY Supplier_Account.TID;"
    RptSuppAcct.Show
End Sub

Private Sub mnRptGE_Click()
    Unload RptExpense
    RptSql = "SELECT * FROM Expenditure ORDER BY TID;"
    RptExpense.Show
End Sub

Private Sub mnRptGEBD_Click()
    Unload RptExpense
    frmDateSelector.Show vbModal
    RptSql = "SELECT * FROM Expenditure WHERE Date BETWEEN '" + RptDate1 + "' AND '" + RptDate2 + "' ORDER BY TID;"
    RptExpense.Show
End Sub





Private Sub mnRptInv_Click()
    ParentForm = "RptInv"
    GridSQLString = "Select Sales.Invoice_No,Sales.Date,Sales.Salesman,Customer.Name,Sales.Grand_Total,Sales.Discount,Sales.Amount_Paid,Sales.Amount_Change,Sales.Amount_Due FROM Sales,Customer WHERE Sales.Customer_ID=Customer.Customer_ID"
    SelectedField = 0
    frmDataSelect.Show vbModal
    
    If ReturnValue = "" Then
        RptStr = InputBox("Please Provide a Invoice No: ", "Information Required")
    Else
        RptStr = ReturnValue
    End If
    
    RptSql = "SELECT Sales.Invoice_No,Sales.Date,Sales.Salesman,Sales.Grand_Total,Sales.Discount,Sales.Payment_Mode,Sales.Amount_Paid,Sales.Amount_Due,Invoice.Product_ID,Stock.Product,Invoice.Quantity,Invoice.Price,Invoice.Net_Total,Customer.Name,Customer.Address,Customer.Phone_No,Customer.Mobile_No FROM Sales,Invoice,Customer,Stock WHERE Sales.Invoice_No='" + RptStr + "' AND Invoice.Invoice_No='" + RptStr + "' AND Invoice.Product_ID=Stock.Product_ID AND Sales.Customer_ID=Customer.Customer_ID"
    RptInvoice.Show
End Sub

Private Sub mnRptOKCus_Click()
    Unload RptCusAcct
    RptSql = "SELECT Customer_Account.TID, Customer_Account.Date, Customer.Name, Customer_Account.Invoice_No, Customer_Account.Total_Amount, Customer_Account.Payment_Mode, Customer_Account.Amount_Paid, Customer_Account.Amount_Due FROM Customer_Account,Customer WHERE Customer_Account.Customer_ID=Customer.Customer_ID AND Customer_Account.Amount_Due=0 ORDER BY Customer_Account.TID;"
    RptCusAcct.Show
End Sub

Private Sub mnRptOKSupp_Click()
    Unload RptSuppAcct
    RptSql = "SELECT Supplier_Account.TID, Supplier_Account.Date,Supplier.Company,Supplier_Account.PO_No,Supplier_Account.Total_Amount,Supplier_Account.Payment_Mode,Supplier_Account.Paid_Amount,Supplier_Account.Due_Amount FROM Supplier_Account,Supplier WHERE Supplier_Account.Supplier_ID=Supplier.Supplier_ID AND Supplier_Account.Due_Amount=0 ORDER BY Supplier_Account.TID;"
    RptSuppAcct.Show
End Sub

Private Sub mnRptPO_Click()
    Unload RptPO
    ParentForm = "RptPO"
    GridSQLString = "Select Purchase_Order.PO_No,Purchase_Order.Date,Purchase_Order.Delivery_Date,Supplier.Company FROM Purchase_Order,Supplier WHERE Purchase_Order.Supplier_ID=Supplier.Supplier_ID"
    SelectedField = 0
    
    frmDataSelect.Show vbModal
    
    If ReturnValue = "" Then
        RptStr = InputBox("Please Provide a Purchase Order No: ", "Information Required")
    Else
        RptStr = ReturnValue
    End If
    
    RptSql = "SELECT Purchase_Order.PO_No,Purchase_Order.Date,Purchase_Order.Delivery_Date,PO_Details.Product,PO_Details.Product_Type,PO_Details.Product_Size,PO_Details.Quantity,PO_Details.Description,Supplier.Company,Supplier.Address,Supplier.Office_No,Supplier.Mobile_No FROM Purchase_Order,PO_Details,Supplier WHERE Purchase_Order.PO_No='" + RptStr + "' AND PO_Details.PO_No='" + RptStr + "' AND Purchase_Order.Supplier_ID=Supplier.Supplier_ID;"
    RptPO.Show
End Sub

Private Sub mnRptPOs_Click()
    Unload RptPO
    RptSql = "SELECT Purchase_Order.PO_No,Purchase_Order.Date,Purchase_Order.Delivery_Date,PO_Details.Product,PO_Details.Product_Type,PO_Details.Product_Size,PO_Details.Quantity,PO_Details.Description,Supplier.Company,Supplier.Address,Supplier.Office_No,Supplier.Mobile_No FROM Purchase_Order,PO_Details,Supplier WHERE Purchase_Order.PO_No=PO_Details.PO_No AND Purchase_Order.Supplier_ID=Supplier.Supplier_ID;"
    RptPO.Show
End Sub

Private Sub mnRptProfit_Click()
    Unload RptProfit
    RptSql = "SELECT * FROM Profit ORDER BY Date,Year,Month"
    RptPathIs = "\Profit.rpt"
    RptProfit.Show
End Sub

Private Sub mnRptProfitChart_Click()
    Unload RptProfit 'Monthly Profit Chart
    
    Dim rpSt As String
    rpSt = InputBox("Please provide the Year for which the report is required!", "Information Required", 2007)
    If rpSt = "" Or numberofcopies = " " Then Exit Sub
    
    RptSql = "SELECT * FROM Profit WHERE Year='" & rpSt & "' ORDER BY Year,Month"
    RptPathIs = "\Profit_1YM.rpt"
    RptProfit.Show
End Sub

Private Sub mnRptProfitYearChart_Click()
    Unload RptProfit 'Yearly Profit Chart
     
    RptSql = "SELECT * FROM Profit ORDER BY Year"
    RptPathIs = "\Profit_Y.rpt"
    RptProfit.Show

End Sub

Private Sub mnRptPS_Click()
    Unload RptPO
    ParentForm = "RptStock"
    GridSQLString = "Select Product_Type from Stock GROUP BY Product_Type;"
    SelectedField = 0
    
    frmDataSelect.Show vbModal
    
    If ReturnValue = "" Then
        RptStr = InputBox("Please Provide Product Type that you are looking for: ", "Information Required")
    Else
        RptStr = ReturnValue
    End If
    
    RptSql = "SELECT Product_ID,Date,Product,Product_Type,Product_Size,Company,Stock_In_Hand,Buying_Price,Selling_Price,ReOrder_Level FROM Stock WHERE Product_Type='" + RptStr + "' ORDER BY Product_ID;"
    RptStock.Show
End Sub

Private Sub mnRptReceiving_Click()
    Unload RptReceivings
    RptSql = "SELECT Receivings.TID,Receivings.Date,Receivings.PO_No,Receivings.Product_ID,Stock.Product,Receivings.Quantity,Receivings.Price,Receivings.Price_per_unit FROM Receivings,Stock  WHERE Receivings.Product_ID=Stock.Product_ID ORder By Receivings.TID;"
    RptReceivings.Show
End Sub

Private Sub mnRptS_Click()
'    DataEnvironment.Commands(1).CommandType = adCmdText
'    DataEnvironment.Commands(1).CommandText = "SELECT SUM(Amount_Paid) FROM Sales"
'    DataEnvironment.Commands(1).Execute
'
'    DataEnvironment.Commands(2).CommandType = adCmdText
'    DataEnvironment.Commands(2).CommandText = "SELECT SUM(Price) FROM G_Sale"
'    DataEnvironment.Commands(2).Execute
'
'    DataEnvironment.Commands(3).CommandType = adCmdText
'    DataEnvironment.Commands(3).CommandText = "SELECT SUM(Amount) FROM Expenditure"
'    DataEnvironment.Commands(3).Execute
'
'    DataEnvironment.Commands(4).CommandType = adCmdText
'    DataEnvironment.Commands(4).CommandText = "SELECT SUM(Amount) FROM Income"
'    DataEnvironment.Commands(4).Execute
'
'    DataEnvironment.Commands(5).CommandType = adCmdText
'    DataEnvironment.Commands(5).CommandText = "SELECT SUM(Total_Bills_Amount),SUM(Total_Due) FROM Supplier"
'    DataEnvironment.Commands(5).Execute
'
'    DataEnvironment.Commands(6).CommandType = adCmdText
'    DataEnvironment.Commands(6).CommandText = "SELECT SUM(Total_Bills_Amount),SUM(Total_Due) FROM Customer"
'    DataEnvironment.Commands(6).Execute
'
'    'UnloadForms
'    If DataEnvironment.rsCusAccount.State = 1 Then
'    DataEnvironment.rsCusAccount.Close
'    End If
'    Summary.Visible = True

    frmSummary.Show
    frmSummary.Left = 20
    frmSummary.Top = 20
End Sub

Private Sub mnRptSales_Click()
    Unload RptSales
    'RptSql = "SELECT Invoice_No,Date,Salesman,Customer_id,Grand_total,Discount,Payment_Mode,Amount_Paid,Amount_Due From Sales Order By Invoice_No;"
    RptSql = "SELECT Sales.Invoice_No,Sales.Date,Sales.Salesman,Sales.Customer_id,Customer.Name,Sales.Grand_total,Sales.Discount,Sales.Payment_Mode,Sales.Amount_Paid,Sales.Amount_Change,Sales.Amount_Due From Sales,Customer WHERE Customer.Customer_ID=Sales.Customer_ID ORDER BY Invoice_No;"
    RptSales.Show
End Sub

Private Sub mnRptSBD_Click()
    Unload RptSales
    
    frmDateSelector.Show vbModal
    
    RptSql = "SELECT Invoice_No,Date,Salesman,Customer_id,Grand_total,Discount,Payment_Mode,Amount_Paid,Amount_change,Amount_Due From Sales WHERE Date BETWEEN '" + RptDate1 + "' AND '" + RptDate2 + "'Order By Invoice_No;"
    RptSales.Show
End Sub

Private Sub mnRptSPOBD_Click()
    Unload RptPO
    
    frmDateSelector.Show vbModal
    RptSql = "SELECT Purchase_Order.PO_No,Purchase_Order.Date,Purchase_Order.Delivery_Date,PO_Details.Product,PO_Details.Product_Type,PO_Details.Product_Size,PO_Details.Quantity,PO_Details.Description,Supplier.Company,Supplier.Address,Supplier.Office_No,Supplier.Mobile_No FROM Purchase_Order,PO_Details,Supplier WHERE Purchase_Order.Date BETWEEN '" + RptDate1 + "' AND '" + RptDate2 + "' AND Purchase_Order.PO_No=PO_Details.PO_No AND Purchase_Order.Supplier_ID=Supplier.Supplier_ID;"
    RptPO.Show
End Sub

Private Sub mnRptSRBD_Click()
    Unload RptReceivings
    frmDateSelector.Show vbModal
    RptSql = "SELECT Receivings.TID,Receivings.Date,Receivings.PO_No,Receivings.Product_ID,Stock.Product,Receivings.Quantity,Receivings.Price,Receivings.Price_per_unit FROM Receivings,Stock WHERE Receivings.Date BETWEEN '" + RptDate1 + "' AND '" + RptDate2 + "' AND Receivings.Product_ID=Stock.Product_ID ORder By Receivings.TID;"
    RptReceivings.Show
End Sub

Private Sub mnRptStock_Click()
    RptSql = "SELECT Product_ID,Date,Product,Product_Type,Product_Size,Company,Stock_In_Hand,Buying_Price,Selling_Price,ReOrder_Level FROM Stock ORDER BY Product_ID;"
    RptStock.Show
End Sub

Private Sub mnRptSupp_Click()
    RptSupplier.Show
End Sub

Private Sub mnRptSuppAcc_Click()
    Unload RptSuppAcct
    RptSql = "SELECT Supplier_Account.TID, Supplier_Account.Date,Supplier.Company,Supplier_Account.PO_No,Supplier_Account.Total_Amount,Supplier_Account.Payment_Mode,Supplier_Account.Paid_Amount,Supplier_Account.Due_Amount FROM Supplier_Account,Supplier WHERE Supplier_Account.Supplier_ID=Supplier.Supplier_ID ORDER BY Supplier_Account.TID;"
    RptSuppAcct.Show
End Sub


Private Sub mnRptTRec_Click()
    Unload RptReceivings
    GetDate
    'RptSql = "SELECT Receivings.TID,Receivings.Date,Receivings.PO_No,Receivings.Product_ID,Stock.Product,Receivings.Quantity,Receivings.Price,Receivings.Price_per_unit FROM Receivings,Stock  WHERE Receivings.Product_ID=Stock.Product_ID AND Receivings.Date='" + DateToday + "'ORDER By Receivings.TID;"
    RptSql = "SELECT Receivings.TID,Receivings.Date,Receivings.PO_No,Receivings.Product_ID,Stock.Product,Receivings.Quantity,Receivings.Price,Receivings.Price_per_unit FROM Receivings,Stock WHERE Receivings.Product_ID=Stock.Product_ID AND Receivings.Date='" + DateToday + "';"
    RptSales.Show
End Sub

Private Sub mnRptTS_Click()
    Unload RptSales
    GetDate
    'RptSql = "SELECT Invoice_No,Date,Salesman,Customer_id,Grand_total,Discount,Payment_Mode,Amount_Paid,Amount_Due From Sales WHERE Date='" + DateToday + "'Order By Invoice_No;"
    RptSql = "SELECT Sales.Invoice_No,Sales.Date,Sales.Salesman,Sales.Customer_id,Customer.Name,Sales.Grand_total,Sales.Discount,Sales.Payment_Mode,Sales.Amount_Paid,Sales.Amount_Change,Sales.Amount_Due From Sales,Customer WHERE Sales.Date='" + DateToday + "' AND Customer.Customer_ID=Sales.Customer_ID ORDER BY Invoice_No;"
    RptSales.Show
End Sub

Private Sub mnSoftware_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnSuppAcc_Click()
    frmSupplierAccount.Show
    frmSupplierAccount.Left = 20
    frmSupplierAccount.Top = 20
End Sub

Private Sub mnSupplier_Click()
    frmSupplier.Show
    frmSupplier.Left = 20
    frmSupplier.Top = 20
End Sub

Private Sub mnUM_Click()
    frmSecurity.Show
    frmSecurity.Left = 20
    frmSecurity.Top = 20
End Sub

Private Sub mnVertical_Click()
    MDIForm1.Arrange vbVertical
End Sub

Private Sub mnArrange_Click()
    MDIForm1.Arrange vbArrangeIcon
End Sub

Private Sub mnCascade_Click()
    MDIForm1.Arrange vbCascade
End Sub

Private Sub mnChangePwd_Click()
    frmChange.Show vbModal
End Sub
