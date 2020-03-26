Attribute VB_Name = "Module"
Option Explicit
Public UserName, Pass, LoginTime, UserTypeUsing, SQLString, SQLErr, GridSQLString, NewSupplier, RptName, Query, RptSql, RptStr, RptPathIs, RptDate1, RptDate2 As String
Public Starting, isStockMinus, isReOrder As Boolean

Public DateToday, DateYear, DateMonth, DateMonthS, DateDay As String
Public TranID, sqlExp, sqlSale, sqlE, sqlFinal As String
Public ExistingExpense, ExistingSale, ExistingActualSale, ExistingProfit, NewExpense, NewSale, NewActualSale, NewProfit, ExpToday, SaleToday, ActualSaleToday, ProfitToday, FinalProfit As Double

Public SelectedField, n, c As Integer
Public ReturnValue As String
Public ParentForm As String

'SendMessage API
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'CB Constants
Public Const CB_MAXLENGTH = 50
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_LIMITTEXT = &H141

'Mouse Cursor
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GCL_HCURSOR = (-12)
Public hOldCursor As Long

'Database
Public conn As ADODB.Connection
Public RsGrid As New ADODB.Recordset
Public RsLogin As New ADODB.Recordset

Public rsSupplier As New ADODB.Recordset
Public RsSuppGrid As New ADODB.Recordset

Public rsSupplierAccount As New ADODB.Recordset
Public RsSuppAccountGrid As New ADODB.Recordset

Public rsPO As New ADODB.Recordset
Public RsPOGrid As New ADODB.Recordset
Public rsPODetails As New ADODB.Recordset
Public RsPODetailsGrid As New ADODB.Recordset

Public rsStock As New ADODB.Recordset
Public rsStockGrid As New ADODB.Recordset

Public rsReceivings As New ADODB.Recordset
Public RsReceivingsGrid As New ADODB.Recordset

Public rsCustomer As New ADODB.Recordset
Public RsCustomerGrid As New ADODB.Recordset

Public rsInvoice As New ADODB.Recordset
Public RsInvoiceGrid As New ADODB.Recordset

Public rsExpense As New ADODB.Recordset
Public RsExpenseGrid As New ADODB.Recordset

Public RsUser As New ADODB.Recordset
Public RsUserGrid As New ADODB.Recordset

Public rsCombo As New ADODB.Recordset
Public rsTmp As New ADODB.Recordset

Private ST As String
Public Rx, RxOS, RxIC, RxNS As Long
Public AddNewStatus As Boolean
Public xCount, xCountIC, xCountOS, xCountNS As Integer
Public db_name, db_server, db_port, db_user, db_pass, constr As String

'TextBox Limit
Public Const ES_NUMBER = &H2000&
Public Const GWL_STYLE = (-16)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Sub Main()
On Error GoTo DBerror

    db_name = "MM_POS"
    db_server = "localhost"
    db_port = ""
    db_user = "root"
    db_pass = "samsung"
    
    Connect
    
    Rx = 0
    RsLogin.Open "SELECT * FROM Login", conn
    RsLogin.Close
    
    frmSplash.Show
    'frmLogin.Show
    'MDIForm1.Show
    
    Exit Sub

DBerror:
    CreateDatabase
End Sub

Public Function Connect()
    constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & db_name & ";SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
    Set conn = New ADODB.Connection
    conn.Open constr
End Function

Private Sub CreateDatabase()
    On Error GoTo ServerErr
        
    'Create...
    conn.Execute "CREATE TABLE Login(User varchar(15), Password varchar(10), Type varchar(15), Name varchar(20), Designation varchar(20), Remarks varchar(50))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Profit(TID varchar(15), Date Date,year varchar(5),month varchar(15), Expense int, Sale int,ActualSale int, Profit int)", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Supplier(Supplier_ID varchar(20) Primary Key, Date Date, Name varchar(30), Company varchar(20), Contact_Person varchar(30), Address varchar(50), Office_No varchar(15), Mobile_No varchar(15), Other_No varchar(15), Fax_No varchar(15), Total_Bills_Amount int, Total_Due int, Remarks varchar(50))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Customer(Customer_ID varchar(20) Primary Key, Date Date, Name varchar(30), CNIC_No varchar(15), Address varchar(50), Occupation varchar(30), Phone_No varchar(15), Mobile_No varchar(15), Other_No varchar(15), Total_Bills_Amount int, Total_Due int, Remarks varchar(50))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Sales(Invoice_No varchar(20) Primary Key, Date Date, Salesman varchar(20), Customer_ID varchar(20), Grand_Total int, Discount varchar(10), Payment_Mode varchar(15), Amount_Paid int, Amount_change int, Amount_Due int,Buying_Total,Profit int, Remarks varchar(50), CONSTRAINT fk_cust_id FOREIGN KEY (Customer_ID) REFERENCES Customer(Customer_ID))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Purchase_Order(PO_No varchar(20) Primary Key, Date Date, Supplier_ID varchar(20), Delivery_Date Date, Remarks varchar(50), CONSTRAINT fk_supp_id1 FOREIGN KEY (Supplier_ID) REFERENCES Supplier(Supplier_ID))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Stock(Product_ID varchar(20) Primary Key, Date Date, Product varchar(30), Product_Type varchar(20), Product_Size varchar(10), Company varchar(30), Stock_In_Hand int, Description varchar(25),Buying_Price int,Selling_Price int, ReOrder_Level int, Remarks varchar(50))", , adExecuteNoRecords
    
    conn.Execute "CREATE TABLE Expenditure(TID varchar(20) Primary Key, Date Date, Expense_Type varchar(20),Supplier varchar(30), Payment_Mode varchar(20),Particulars varchar(30),Amount int,Remarks varchar(50))", , adExecuteNoRecords
    
    conn.Execute "CREATE TABLE PO_Details(TID varchar(20) Primary Key, PO_No varchar(20), Product varchar(20), Product_Type varchar(20), Product_Size varchar(10), Quantity int, Description varchar(30), CONSTRAINT fk_PO1_No FOREIGN KEY (PO_No) REFERENCES Purchase_Order(PO_No))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Supplier_Account(TID varchar(20) Primary Key, Supplier_ID varchar(20), Date Date, PO_No varchar(20), Total_Amount int, Payment_Mode varchar(15), Paid_Amount int, Due_Amount int, Remarks varchar(50), CONSTRAINT fk_supp_id2 FOREIGN KEY (Supplier_ID) REFERENCES Supplier(Supplier_ID), CONSTRAINT fk_PO_No FOREIGN KEY (PO_No) REFERENCES Purchase_Order(PO_No))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Invoice(TID varchar(20) Primary Key, Invoice_No varchar(20), Product_ID varchar(20), Quantity int, Price int, Net_Total int, CONSTRAINT fk_Inv_No FOREIGN KEY (Invoice_No) REFERENCES Sales(Invoice_No), CONSTRAINT fk_prod_id FOREIGN KEY (Product_ID) REFERENCES Stock(Product_ID))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Customer_Account(TID varchar(20) Primary Key, Customer_ID varchar(20), Date Date, Invoice_No varchar(20), Total_Amount int, Payment_Mode varchar(15), Amount_Paid int, Amount_Due int, Remarks varchar(50), CONSTRAINT fk_Inv1_No FOREIGN KEY (Invoice_No) REFERENCES Sales(Invoice_No))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Receivings(TID varchar(20) Primary Key, Date Date, PO_No varchar(20), Product_ID varchar(20), Quantity int, Price int, Price_per_unit int, Remarks varchar(50), CONSTRAINT fk_PO_No1 FOREIGN KEY (PO_No) REFERENCES Purchase_Order(PO_No), CONSTRAINT fk_prod_id3 FOREIGN KEY (Product_ID) REFERENCES Stock(Product_ID))", , adExecuteNoRecords
    
    'Insert...
    conn.Execute "INSERT INTO Login values('admin','admin','Admin','Admin User','Administration','-')", , adExecuteNoRecords
    conn.Execute "INSERT INTO Login values('manager','manager','Manager','Manager User','Management','-')", , adExecuteNoRecords
    conn.Execute "INSERT INTO Login values('salesman','salesman','Salesman','Sales User','Sales Dept','-')", , adExecuteNoRecords
    
    MsgBox "DATABASE CREATED & POPULATED WITH INITIAL DATA, Please RUN PROGRAM AGAIN!!!", vbInformation, "POS"
    Exit Sub
ServerErr:
    MsgBox "Unable to Locate Database On Server, Please RUN PROGRAM AGAIN!!!", vbInformation, "DATABASE NOT FOUND"
    constr = "Provider=MSDASQL.1;Password=samsung;Persist Security Info=True;User ID=root;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 3.51 Driver};DESC=;SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
    Set conn = New ADODB.Connection
    conn.Open constr
    conn.Execute "Create Database " & Trim$("MM_POS"), , adExecuteNoRecords
    End
End Sub

Public Sub GetDate()
    DateYear = Year(Date)
    DateMonth = Month(Date)
    DateDay = Day(Date)
    
    DateToday = "" & DateYear & "-" & DateMonth & "-" & DateDay
    
    If DateMonth = 1 Then DateMonthS = "01-January"
    If DateMonth = 2 Then DateMonthS = "02-February"
    If DateMonth = 3 Then DateMonthS = "03-March"
    If DateMonth = 4 Then DateMonthS = "04-April"
    If DateMonth = 5 Then DateMonthS = "05-May"
    If DateMonth = 6 Then DateMonthS = "06-June"
    If DateMonth = 7 Then DateMonthS = "07-July"
    If DateMonth = 8 Then DateMonthS = "08-August"
    If DateMonth = 9 Then DateMonthS = "09-September"
    If DateMonth = 10 Then DateMonthS = "10-October"
    If DateMonth = 11 Then DateMonthS = "11-November"
    If DateMonth = 12 Then DateMonthS = "12-December"
End Sub

'Combo
Public Sub Combo_Lookup(ctlCombo As ComboBox)
   Dim lngItemPos As Long
   Dim strCombo As String

   strCombo = ctlCombo.Text

   ' Use SendMessage() API to Find Combobox Values
   lngItemPos = SendMessage(ctlCombo.hwnd, CB_FINDSTRING, -1, ByVal strCombo)

   If lngItemPos >= 0 Then
      ctlCombo.ListIndex = lngItemPos
   End If

   ctlCombo.SelStart = Len(strCombo)
   ctlCombo.SelLength = CB_MAXLENGTH
End Sub

Public Sub UnloadForms()
    'Unload frmSecurity
    Unload frmSupplier
End Sub

Public Sub UpdateProfit()
    
    GetDate
    TranID = "T" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
        
    sqlExp = "SELECT SUM(Amount) as 'Sum' FROM Expenditure WHERE Date='" & DateToday & "';" 'Getting Expense Total
    GetExpenseData (sqlExp)
        
    sqlSale = "SELECT SUM(Grand_Total) as 'Sale',SUM(Buying_Total) as 'ActualSale',SUM(Profit) as 'Profit' FROM Sales WHERE Date='" & DateToday & "';" 'Getting Total Sale, ActualSale and profit
    GetSaleData (sqlSale)
        
    sqlE = "SELECT Expense,Sale,ActualSale,Profit FROM Profit WHERE Date='" & DateToday & "';" 'Getting Existing Values
    GetExistingData (sqlE)
    
    FinalProfit = ProfitToday - ExpToday
    
    If DupCheck("SELECT * from Profit WHERE Date='" & DateToday & "' AND Year='" & DateYear & "' AND Month='" & DateMonthS & "'") = True Then
        'if already exists
        sqlFinal = "UPDATE Profit SET Expense=" & ExpToday & ",Sale=" & SaleToday & ",ActualSale=" & ActualSaleToday & ",Profit=" & FinalProfit & " WHERE Date='" & DateToday & "';"
    Else
        If (ExpToday = 0 And SaleToday = 0 And FinalProfit = 0) Then
            Exit Sub
        Else
            sqlFinal = "INSERT INTO Profit VALUES('" & TranID & "','" & DateToday & "','" & DateYear & "','" & DateMonthS & "'," & ExpToday & "," & SaleToday & "," & ActualSaleToday & "," & ProfitToday & ")"
        End If
    End If
    
    conn.Execute sqlFinal

End Sub
Private Function DupCheck(chkID As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    rs.Open chkID, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        DupCheck = False
        Exit Function
    Else
        DupCheck = True
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Sub GetExpenseData(sqlED As String)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    rs.Open sqlED, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        ExpToday = 0
        Exit Sub
    End If
    If rs!Sum <> "" Then
        ExpToday = rs!Sum
    Else
        ExpToday = 0
    End If
    rs.Close
    Set rs = Nothing
End Sub
Public Sub GetSaleData(sqlSD As String)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    rs.Open sqlSD, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        SaleToday = 0
        ProfitToday = 0
        ActualSaleToday = 0
        Exit Sub
    End If
    If rs!Sale <> "" Then
        SaleToday = rs!Sale
        ProfitToday = rs!Profit
        ActualSaleToday = rs!ActualSale
    Else
        SaleToday = 0
        ProfitToday = 0
        ActualSaleToday = 0
    End If
    rs.Close
    Set rs = Nothing
End Sub
Public Sub GetExistingData(sqlED As String)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    rs.Open sqlED, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        ExistingExpense = 0
        ExistingSale = 0
        ExistingProfit = 0
        ExistingActualSale = 0
        Exit Sub
    End If
    If rs!Expense <> "" Then
        ExistingExpense = rs!Expense
        ExistingSale = rs!Sale
        ExistingProfit = rs!Profit
        ExistingActualSale = rs!ActualSale
    Else
        ExistingExpense = 0
        ExistingSale = 0
        ExistingProfit = 0
        ExistingActualSale = 0
    End If
    rs.Close
    Set rs = Nothing
End Sub

Public Sub BackupDatabase()
    On Error GoTo Error
    'Shell "C:\Program Files\Inventory\Download.bat", vbNormalFocus
    Shell App.Path + "\Download.bat", vbNormalFocus
    Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub


'Table backup technique.
'SELECT * INTO OUTFILE 'C:\Backup\Backup.adl' FIELDS TERMINATED BY ',' LINES TERMINATED BY '\n' FROM TableName;"

'LOAD DATA INFILE '" + path +     "' INTO TABLE " +     CboTableRst.getSelectedItem() +    " FIELDS TERMINATED BY ',' LINES TERMINATED BY '\\n';");

Public Sub CheckUser()
    If (UserTypeUsing = "Manager") Then
        MDIForm1.mnPurchase.Enabled = False
        MDIForm1.mnSM.Enabled = False
        MDIForm1.mnSales.Enabled = False
        MDIForm1.mnUM.Enabled = False
        MDIForm1.mnGeneral.Enabled = False
        MDIForm1.mnCB.Enabled = False
    End If
    If (UserTypeUsing = "Salesman") Then
        MDIForm1.mnPurchase.Enabled = False
        MDIForm1.mnSM.Enabled = False
        MDIForm1.mnUM.Enabled = False
        MDIForm1.mnGeneral.Enabled = False
        MDIForm1.mnSupplierR.Enabled = False
        MDIForm1.mnRptPr.Enabled = False
        MDIForm1.mnSaleR.Enabled = False
        MDIForm1.mnProfit.Enabled = False
        MDIForm1.mnRptS.Enabled = False
        MDIForm1.mnRptG.Enabled = False
    End If
    If (UserTypeUsing = "Stock Manager") Then
        MDIForm1.mnUM.Enabled = False
    End If
End Sub
