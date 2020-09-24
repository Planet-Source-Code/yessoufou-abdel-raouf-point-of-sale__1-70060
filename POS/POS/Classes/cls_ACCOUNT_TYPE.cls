VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_ACCOUNT_TYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarAccountType As String 'local copy
Public Property Let AccountType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountType = 5
    mvarAccountType = vData
End Property


Public Property Get AccountType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountType
    AccountType = mvarAccountType
End Property


Public Function fn_LOAD_ACCOUNT_TYPE(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        cls_DATABASE_Obj.strQueryString = "SELECT * FROM tb_ACCOUNT_TYPE"
        Else
            cls_DATABASE_Obj.strQueryString = "SELECT * FROM tb_ACCOUNT_TYPE WHERE AccountTypeID = " & lngID
    End If

    Set rec = cls_DATABASE_Obj.fn_MANAGE_RECORD

    Set fn_LOAD_ACCOUNT_TYPE = rec

End Function

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

Public Function fn_SAVE_ACCOUNT_TYPE_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_ACCOUNT_TYPE " _
    & "(AccountType) " _
    & " VALUES " _
    & "('" & mvarAccountType & "')"

    Set fn_SAVE_ACCOUNT_TYPE_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_ACCOUNT_TYPE_RECORDS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_ACCOUNT_TYPE SET " _
    & "AccountType = '" & mvarAccountType & "'" _
    & " WHERE AccountTypeID = " & lngID
    
    Set fn_UPDATE_ACCOUNT_TYPE_RECORDS = db.fn_MANAGE_RECORD

End Function

