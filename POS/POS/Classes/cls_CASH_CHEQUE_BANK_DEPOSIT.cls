VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_CASH_CHEQUE_BANK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarTransactionID As Long 'local copy
Private mvarTransactionName As String 'local copy
Private mvarAgencyID As Long 'local copy
Private mvarAgencyName As String 'local copy
Private mvarChequeNo As String 'local copy
Private mvarBankID As Long 'local copy
Private mvarBankName As String 'local copy
Private mvarAccountNo As String 'local copy
Private mvarChequeDate As Date 'local copy
Private mvarAmount As Double 'local copy
Private mvarDateReceived As Date 'local copy
Private mvarPaidBy As String 'local copy
Private mvarChequeStatus As String 'local copy
Private mvarBankDeposited As String 'local copy
Private mvarBranchCode As String 'local copy
Private mvarTransactionNo As String 'local copy
Private mvarCurrencyName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCustomerID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarCustomerName As String 'local copy
Public Property Let CustomerName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomerName = 5
    mvarCustomerName = vData
End Property


Public Property Get CustomerName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CustomerName
    CustomerName = mvarCustomerName
End Property



Public Property Let CustomerID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomerID = 5
    mvarCustomerID = vData
End Property


Public Property Get CustomerID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CustomerID
    CustomerID = mvarCustomerID
End Property



Public Property Let CurrencyName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrencyName = 5
    mvarCurrencyName = vData
End Property


Public Property Get CurrencyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrencyName
    CurrencyName = mvarCurrencyName
End Property



Public Property Let TransactionNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TransactionNo = 5
    mvarTransactionNo = vData
End Property


Public Property Get TransactionNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TransactionNo
    TransactionNo = mvarTransactionNo
End Property



Public Property Let BranchCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BranchCode = 5
    mvarBranchCode = vData
End Property


Public Property Get BranchCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BranchCode
    BranchCode = mvarBranchCode
End Property



Public Property Let BankDeposited(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BankDeposited = 5
    mvarBankDeposited = vData
End Property


Public Property Get BankDeposited() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BankDeposited
    BankDeposited = mvarBankDeposited
End Property



Public Property Let ChequeStatus(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChequeStatus = 5
    mvarChequeStatus = vData
End Property


Public Property Get ChequeStatus() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChequeStatus
    ChequeStatus = mvarChequeStatus
End Property



Public Property Let PaidBy(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PaidBy = 5
    mvarPaidBy = vData
End Property


Public Property Get PaidBy() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PaidBy
    PaidBy = mvarPaidBy
End Property



Public Property Let DateReceived(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateReceived = 5
    mvarDateReceived = vData
End Property


Public Property Get DateReceived() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateReceived
    DateReceived = mvarDateReceived
End Property



Public Property Let Amount(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    mvarAmount = vData
End Property


Public Property Get Amount() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Amount
    Amount = mvarAmount
End Property



Public Property Let ChequeDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChequeDate = 5
    mvarChequeDate = vData
End Property


Public Property Get ChequeDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChequeDate
    ChequeDate = mvarChequeDate
End Property



Public Property Let AccountNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountNo = 5
    mvarAccountNo = vData
End Property


Public Property Get AccountNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountNo
    AccountNo = mvarAccountNo
End Property



Public Property Let BankName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BankName = 5
    mvarBankName = vData
End Property


Public Property Get BankName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BankName
    BankName = mvarBankName
End Property



Public Property Let BankID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BankID = 5
    mvarBankID = vData
End Property


Public Property Get BankID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BankID
    BankID = mvarBankID
End Property



Public Property Let ChequeNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChequeNo = 5
    mvarChequeNo = vData
End Property


Public Property Get ChequeNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChequeNo
    ChequeNo = mvarChequeNo
End Property



Public Property Let AgencyName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AgencyName = 5
    mvarAgencyName = vData
End Property


Public Property Get AgencyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AgencyName
    AgencyName = mvarAgencyName
End Property



Public Property Let AgencyID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AgencyID = 5
    mvarAgencyID = vData
End Property


Public Property Get AgencyID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AgencyID
    AgencyID = mvarAgencyID
End Property



Public Property Let TransactionName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TransactionName = 5
    mvarTransactionName = vData
End Property


Public Property Get TransactionName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TransactionName
    TransactionName = mvarTransactionName
End Property



Public Property Let TransactionID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TransactionID = 5
    mvarTransactionID = vData
End Property


Public Property Get TransactionID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TransactionID
    TransactionID = mvarTransactionID
End Property

Private Sub Class_Initialize()
'On Error GoTo errHandler

    Set con_Obj = mdl_MAIN.fn_OPEN_CONNECTION

EXITPROCEDURE:
    Exit Sub
    
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call Mdl_Functions.fnWriteErrorToFile(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
'    GoTo EXITPROCEDURE
End Sub

Private Sub Class_Terminate()
On Error GoTo errHandler

    Call mdl_MAIN.fn_CLOSE_CONNECTION(con_Obj)

EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Public Function fn_LOAD_CASH_CHEQUE_BANK_DEPOSIT_REP(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    
    Select Case lngID
        Case 0
            cls_DATABASE_Obj.strQueryString = "SELECT * FROM vw_CASH"
        Case 1
            cls_DATABASE_Obj.strQueryString = "SELECT * FROM vw_CHEQUE"
        Case 2
            cls_DATABASE_Obj.strQueryString = "SELECT * FROM vw_BANK_DEPOSIT"
    End Select
    
    Set rec = cls_DATABASE_Obj.fn_MANAGE_RECORD

    Set fn_LOAD_CASH_CHEQUE_BANK_DEPOSIT_REP = rec

End Function


Public Function fn_SAVE_DEPOSIT_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_CUSTOMERS_DEPOSIT " _
    & "(TransactionID,TransactionName,CustomerID,CustomerName,ChequeNo,BankID,BankName,AccountNo,ChequeDate,Amount,DateReceived,PaidBy,ChequeStatus,BankDeposited,BranchCode,TransactionNo) " _
    & " VALUES " _
    & "(" & mvarTransactionID & "," _
    & "'" & mvarTransactionName & "'," _
    & mvarCustomerID & "," _
    & "'" & mvarCustomerName & "'," _
    & "'" & mvarChequeNo & "'," _
    & mvarBankID & "," _
    & "'" & mvarBankName & "'," _
    & "'" & mvarAccountNo & "'," _
    & "'" & mvarChequeDate & "'," _
    & mvarAmount & "," _
    & "'" & mvarDateReceived & "'," _
    & "'" & mvarPaidBy & "'," _
    & "'" & mvarChequeStatus & "'," _
    & "'" & mvarBankDeposited & "'," _
    & "'" & mvarBranchCode & "'," _
    & "'" & mvarTransactionNo & "')"

    Set fn_SAVE_DEPOSIT_RECORDS = db.fn_MANAGE_RECORD

End Function



Public Function fn_ADD_CUSTOMER_AMOUNT(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_CUSTOMERS SET " _
    & "Amount = Amount + " & mvarAmount _
    & " WHERE CustomerID = " & lngID
    
    Set fn_ADD_CUSTOMER_AMOUNT = db.fn_MANAGE_RECORD

End Function

Public Function fn_REDUCE_CUSTOMER_AMOUNT(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_CUSTOMERS SET " _
    & "Amount = " & mvarAmount _
    & " WHERE CustomerID = " & lngID
    
    Set fn_REDUCE_CUSTOMER_AMOUNT = db.fn_MANAGE_RECORD

End Function

