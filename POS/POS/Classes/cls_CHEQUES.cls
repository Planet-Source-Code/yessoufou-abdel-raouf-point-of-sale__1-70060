VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_CHEQUES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
Private mvarChequeDate As Date 'local copy
Private mvarChequeNo As String 'local copy
Private mvarBankID As Long 'local copy
Private mvarAmount As Double 'local copy
Private mvarStatus As Long 'local copy
Private mvarLocation As Long 'local copy
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



Public Property Let Location(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vData
End Property


Public Property Get Location() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
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

Public Function fn_SAVE_CHEQUES() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_CHEQUES " _
    & "(ChequeDate,ChequeNo,BankID,Amount,Status,LocationID,UserID) " _
    & " VALUES " _
    & "('" & mvarChequeDate & "'," _
    & "'" & mvarChequeNo & "'," _
    & mvarBankID & "," _
    & mvarAmount & "," _
    & mvarStatus & "," _
    & mvarLocation & "," _
    & mvarUserID & ")"

    Set fn_SAVE_CHEQUES = db.fn_MANAGE_RECORD

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

