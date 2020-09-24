VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_CASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarCashID As Long 'local copy
Private mvarAmount As Double 'local copy
Private mvarDatePaid As Date 'local copy
Private mvarPaidBy As String 'local copy
'local variable(s) to hold property value(s)
Private mvarAgencyID As Long 'local copy
Private mvarCustomerID As Long 'local copy
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



Public Property Let DatePaid(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DatePaid = 5
    mvarDatePaid = vData
End Property


Public Property Get DatePaid() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DatePaid
    DatePaid = mvarDatePaid
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



Public Property Let CashID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CashID = 5
    mvarCashID = vData
End Property


Public Property Get CashID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CashID
    CashID = mvarCashID
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

Public Function fn_SAVE_CASH_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_CASH " _
    & "(CustomerID,Amount,DatePaid,PaidBy) " _
    & " VALUES " _
    & "(" & mvarCustomerID & "," _
    & mvarAmount & "," _
    & "'" & mvarDatePaid & "'," _
    & "'" & mvarPaidBy & "')"

    Set fn_SAVE_CASH_RECORDS = db.fn_MANAGE_RECORD

End Function

