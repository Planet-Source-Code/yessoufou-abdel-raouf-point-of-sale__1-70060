VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_BANK_TRANSACTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarTransactionDate As Date 'local copy
Private mvarTransactionNo As String 'local copy
Private mvarTransactionType As String 'local copy
Private mvarDescription As String 'local copy
Private mvarAccountID As Long 'local copy
Private mvarAmount As Double 'local copy
Private mvarTransactedBy As String 'local copy
Private mvarPosted As Long 'local copy
Private mvarChequeID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarUserID As Long 'local copy
Public Property Let UserID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserID = 5
    mvarUserID = vData
End Property


Public Property Get UserID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserID
    UserID = mvarUserID
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



Public Property Let Posted(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Posted = 5
    mvarPosted = vData
End Property


Public Property Get Posted() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Posted
    Posted = mvarPosted
End Property



Public Property Let TransactedBy(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TransactedBy = 5
    mvarTransactedBy = vData
End Property


Public Property Get TransactedBy() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TransactedBy
    TransactedBy = mvarTransactedBy
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



Public Property Let AccountID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountID = 5
    mvarAccountID = vData
End Property


Public Property Get AccountID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountID
    AccountID = mvarAccountID
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let TransactionType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TransactionType = 5
    mvarTransactionType = vData
End Property


Public Property Get TransactionType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TransactionType
    TransactionType = mvarTransactionType
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



Public Property Let TransactionDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TransactionDate = 5
    mvarTransactionDate = vData
End Property


Public Property Get TransactionDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TransactionDate
    TransactionDate = mvarTransactionDate
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

Public Function fn_LOAD_BANK_TRANSACTIONS(Optional lngAccountID As Long, Optional fromDate As Date, Optional toDate As Date) As ADODB.Recordset

'    If lngID = 0 Then
'        db.strQueryString = "SELECT * FROM tb_BANK_TRANSACTIONS"
'        Else
    db.strQueryString = "SELECT * FROM tb_BANK_TRANSACTIONS WHERE AccountID = " & lngAccountID & " AND TransactionDate BETWEEN '" & fromDate & "' AND '" & toDate & "'"
'    End If

    Set fn_LOAD_BANK_TRANSACTIONS = db.fn_MANAGE_RECORD


End Function
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

Public Function fn_SAVE_BANK_TRANSACTIONS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_BANK_TRANSACTIONS " _
    & "(TransactionDate,EntryDate,TransactionNo,TransactionType,Description,AccountID,Amount,TransactedBy,Posted,ChequeID,UserID) " _
    & " VALUES " _
    & "('" & mvarTransactionDate & "'," _
    & "'" & Now & "'," _
    & "'" & fn_AUTOGEN & "'," _
    & "'" & mvarTransactionType & "'," _
    & "'" & mvarDescription & "'," _
    & mvarAccountID & "," _
    & mvarAmount & "," _
    & "'" & mvarTransactedBy & "'," _
    & mvarPosted & "," _
    & mvarChequeID & "," _
    & "'" & mvarUserID & "')"

    Set fn_SAVE_BANK_TRANSACTIONS = db.fn_MANAGE_RECORD

End Function

Public Function fn_DELETE_BANK_RANSACTIONS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "DELETE FROM tb_BANK_TRANSACTIONS WHERE TransactionID = " & lngID

    Set fn_DELETE_BANK_RANSACTIONS = db.fn_MANAGE_RECORD

End Function

Public Function fn_POST_BANK_RANSACTIONS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_BANK_TRANSACTIONS SET POSTED = 1 WHERE TransactionID = " & lngID

    Set fn_POST_BANK_RANSACTIONS = db.fn_MANAGE_RECORD

End Function


Public Function fn_AUTOGEN() As String

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(TransactionNo) FROM tb_BANK_TRANSACTIONS"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN = "T00001"
        Else
            fn_AUTOGEN = "T" & Format(Right(Trim(rec(0)), 5) + 1, "00000")
        End If
    Else
        fn_AUTOGEN = "T00001"
    End If

End Function
