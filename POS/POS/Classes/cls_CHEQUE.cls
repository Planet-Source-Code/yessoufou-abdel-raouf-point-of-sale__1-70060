VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_CHEQUE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
Private mvarChequeID As Long 'local copy
Private mvarAgencyName As Long 'local copy
Private mvarChequeNo As String 'local copy
Private mvarBankName As Long 'local copy
Private mvarAccountNo As String  'local copy
Private mvarAmount As Double 'local copy
'Private mvarChequeType As String 'local copy
Private mvarChequeDate As Date 'local copy
Private mvarDateReceived As Date 'local copy
Private mvarPaidBy As String 'local copy
Private mvarStatus As Long 'local copy
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



Public Property Let Status(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
End Property


Public Property Get Status() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
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



'Public Property Let ChequeType(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.ChequeType = 5
'    mvarChequeType = vData
'End Property
'
'
'Public Property Get ChequeType() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.ChequeType
'    ChequeType = mvarChequeType
'End Property



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



Public Property Let BankName(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BankName = 5
    mvarBankName = vData
End Property


Public Property Get BankName() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BankName
    BankName = mvarBankName
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



Public Property Let AgencyName(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AgencyName = 5
    mvarAgencyName = vData
End Property


Public Property Get AgencyName() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AgencyName
    AgencyName = mvarAgencyName
End Property



Public Property Let ChequeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChequeID = 5
    mvarChequeID = vData
End Property


Public Property Get ChequeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChequeID
    ChequeID = mvarChequeID
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

Public Function fn_SAVE_CHECK_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_CHEQUE " _
    & "(CustomerID,ChequeNo,BankName,AccountNo,Amount,ChequeDate,DateReceived,PaidBy,Status) " _
    & " VALUES " _
    & "(" & mvarCustomerID & "," _
    & "'" & mvarChequeNo & "'," _
    & mvarBankName & "," _
    & "'" & mvarAccountNo & "'," _
    & mvarAmount & "," _
    & "'" & mvarChequeDate & "'," _
    & "'" & mvarDateReceived & "'," _
    & "'" & mvarPaidBy & "'," _
    & mvarStatus & ")"

    Set fn_SAVE_CHECK_RECORDS = db.fn_MANAGE_RECORD

End Function

