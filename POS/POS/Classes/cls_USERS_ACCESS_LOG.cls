VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_USERS_ACCESS_LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
Private mvarAccessLogID As Long 'local copy
Private mvarUserID As Long 'local copy
Private mvarWorkStationName As String 'local copy
Private mvarLoginDate As Date  'local copy
Private mvarLoginType As Long 'local copy
Private mvarDescription As String 'local copy
'local variable(s) to hold property value(s)
Private mvarLoginTime As Date  'local copy
Public Property Let LoginTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LoginTime = 5
    mvarLoginTime = vData
End Property


Public Property Get LoginTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LoginTime
    LoginTime = mvarLoginTime
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



Public Property Let LoginType(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LoginType = 5
    mvarLoginType = vData
End Property


Public Property Get LoginType() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LoginType
    LoginType = mvarLoginType
End Property



Public Property Let LoginDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LoginDate = 5
    mvarLoginDate = vData
End Property


Public Property Get LoginDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LoginDate
    LoginDate = mvarLoginDate
End Property



Public Property Let WorkStationName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WorkStationName = 5
    mvarWorkStationName = vData
End Property


Public Property Get WorkStationName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WorkStationName
    WorkStationName = mvarWorkStationName
End Property



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



Public Property Let AccessLogID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccessLogID = 5
    mvarAccessLogID = vData
End Property


Public Property Get AccessLogID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccessLogID
    AccessLogID = mvarAccessLogID
End Property


Private Sub Class_Initialize()
'On Error GoTo errHandler

    Set con_Obj = mdl_MAIN.fn_OPEN_CONNECTION

EXITPROCEDURE:
    Exit Sub
    
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call MdlFunctions.fnWriteErrorToFile(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
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


Public Function fn_LOAD_USERS_ACCESS_LOGS(Optional lngUserID As Long, Optional fromDate As Date, Optional toDate As Date) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngUserID = 0 Then
        db.strQueryString = "SELECT * FROM vw_USERS_ACCESS_LOGS "
        Else
            db.strQueryString = "SELECT * FROM vw_USERS_ACCESS_LOGS WHERE LoginDate BETWEEN '" & fromDate & "' AND '" & toDate & "'"
    End If
    Set rec = db.fn_MANAGE_RECORD
    Set fn_LOAD_USERS_ACCESS_LOGS = rec
    
End Function



Public Function fn_SAVE_USERS_ACCESS_LOG()

    db.strQueryString = "INSERT INTO tb_USERS_ACCESS_LOG " _
    & " (AccessLogID,UserID,WorkStationName,LoginDate,LoginTime,LoginType,Description)" _
    & " VALUES (" _
    & mvarAccessLogID & "," _
    & mvarUserID & "," _
    & "'" & mvarWorkStationName & "'," _
    & "'" & mvarLoginDate & "'," _
    & "'" & mvarLoginTime & "'," _
    & mvarLoginType & "," _
    & "'" & mvarDescription & "')"
    
    fn_SAVE_USERS_ACCESS_LOG = db.fn_MANAGE_RECORD
    
End Function


Public Function fn_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(AccessLogID) FROM tb_USERS_ACCESS_LOG"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN = 1
        Else
            fn_AUTOGEN = Val(rec(0)) + 1
        End If
    Else
        fn_AUTOGEN = 1
    End If

End Function

